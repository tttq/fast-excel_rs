//! `#[derive(ExcelRow)]`：声明式列映射（对标 EasyExcel 的 `@ExcelProperty`）
//!
//! ```ignore
//! #[derive(ExcelRow)]
//! #[excel(sheet = "产品导入")]
//! struct ProductRow {
//!     #[excel(header = "工厂名称", alias = ["工厂", "厂名"], required)]
//!     factory_name: String,
//!
//!     #[excel(header = "报价(USD)", kind = "decimal", format = "#,##0.00")]
//!     price: Option<Decimal>,
//!
//!     #[excel(header = "出货时间(天)")]
//!     lead_time: Option<i32>,
//!
//!     #[excel(header = "图片1", image)]
//!     image1: Option<CellImage>,
//!
//!     /// 未声明列全部收下（动态适配）
//!     #[excel(dynamic)]
//!     extra: HashMap<String, CellValue>,
//! }
//! ```
//!
//! 支持的属性：
//!
//! | 属性 | 位置 | 说明 |
//! |---|---|---|
//! | `sheet = "名"` | 结构体 | 默认工作表名（多 sheet 导出用） |
//! | `header = "列名"` | 字段 | 标准表头；不写则用字段名 |
//! | `alias = "x"` / `alias = ["x","y"]` | 字段 | 别名（动态表头适配） |
//! | `required` | 字段 | 必填（缺列直接报错） |
//! | `kind = "text/integer/number/decimal/bool/date/datetime/image/any"` | 字段 | 语义类型（不写按 Rust 类型推断） |
//! | `width = 16.0` / `format = "#,##0.00"` | 字段 | 导出列宽 / 单元格格式 |
//! | `text` | 字段 | 按文本格式写（长数字、手机号） |
//! | `image` | 字段 | 图片列（导入读单元格内嵌图片） |
//! | `default = "x"` / `note = "x"` | 字段 | 模板默认值 / 说明 |
//! | `dropdown = ["a","b"]` | 字段 | 模板下拉选项 |
//! | `ignore` | 字段 | 不参与导入导出（默认值填充） |
//! | `dynamic` | 字段 | 兜住所有未声明列，类型用 `HashMap<String, CellValue>` 或 `Vec<(String, CellValue)>` |

use proc_macro::TokenStream;
use quote::quote;
use syn::{Data, DeriveInput, Field, Fields, Lit, LitStr, Token, Type, parse_macro_input};

#[proc_macro_derive(ExcelRow, attributes(excel))]
pub fn derive_excel_row(input: TokenStream) -> TokenStream {
    let input = parse_macro_input!(input as DeriveInput);
    match expand(&input) {
        Ok(tokens) => tokens.into(),
        Err(err) => err.to_compile_error().into(),
    }
}

#[derive(Default)]
struct FieldAttrs {
    header: Option<String>,
    aliases: Vec<String>,
    required: bool,
    kind: Option<String>,
    width: Option<f64>,
    format: Option<String>,
    text_format: bool,
    image: bool,
    ignore: bool,
    default: Option<String>,
    note: Option<String>,
    dropdown: Vec<String>,
    dynamic: bool,
    /// 显式指定了 required（否则按 Option 推断）
    required_set: bool,
}

fn expand(input: &DeriveInput) -> syn::Result<proc_macro2::TokenStream> {
    let name = &input.ident;
    let data = match &input.data {
        Data::Struct(data) => data,
        _ => {
            return Err(syn::Error::new_spanned(
                input,
                "ExcelRow 只能派生在结构体上",
            ));
        }
    };
    let fields = match &data.fields {
        Fields::Named(named) => named.named.iter().collect::<Vec<_>>(),
        _ => {
            return Err(syn::Error::new_spanned(
                input,
                "ExcelRow 需要具名字段的结构体",
            ));
        }
    };

    let mut sheet_name: Option<String> = None;
    for attr in &input.attrs {
        if !attr.path().is_ident("excel") {
            continue;
        }
        attr.parse_nested_meta(|meta| {
            if meta.path.is_ident("sheet") {
                sheet_name = Some(meta.value()?.parse::<LitStr>()?.value());
                Ok(())
            } else if meta.path.is_ident("register") {
                // 由 ExcelExecutor 派生处理；ExcelRow 这里忽略即可（两者常组合使用）
                meta.value()?.parse::<LitStr>()?;
                Ok(())
            } else {
                Err(meta.error("结构体上只支持 #[excel(sheet = \"...\", register = \"...\")]"))
            }
        })?;
    }

    let mut column_defs = Vec::new();
    let mut extract_stmts = Vec::new();
    let mut to_row_exprs = Vec::new();
    let mut ctor_fields = Vec::new();
    let mut has_dynamic = false;

    for (index, field) in fields.iter().enumerate() {
        let ident = field
            .ident
            .clone()
            .ok_or_else(|| syn::Error::new_spanned(field, "字段必须有名字"))?;
        let attrs = parse_field_attrs(field)?;

        if attrs.ignore {
            ctor_fields.push(quote! {
                #ident: ::core::default::Default::default()
            });
            continue;
        }
        if attrs.dynamic {
            has_dynamic = true;
            ctor_fields.push(quote! {
                #ident: __dynamic_values.into_iter().collect()
            });
            continue;
        }

        // 表头：显式 header 优先，否则用字段名
        let header = attrs.header.clone().unwrap_or_else(|| ident.to_string());
        let ty = &field.ty;

        // ── columns() ──
        let aliases = attrs.aliases.clone();
        let required = attrs.required;
        let kind = resolve_kind(&attrs, ty);
        let kind_ident = syn::Ident::new(kind, proc_macro2::Span::call_site());
        let width = attrs.width;
        let format = attrs.format.clone();
        let text_format = attrs.text_format;
        let image = attrs.image;
        let default = attrs.default.clone();
        let note = attrs.note.clone();
        let dropdown = attrs.dropdown.clone();
        let header_lit = header.clone();

        let width_stmt = match width {
            Some(w) => quote! { __col = __col.width(#w); },
            None => quote! {},
        };
        let format_stmt = match format {
            Some(f) => quote! { __col = __col.format(#f); },
            None => quote! {},
        };
        let default_stmt = match default {
            Some(d) => quote! { __col = __col.default_value(#d); },
            None => quote! {},
        };
        let note_stmt = match note {
            Some(n) => quote! { __col = __col.note(#n); },
            None => quote! {},
        };
        let required_stmt = if required {
            quote! { __col = __col.required(); }
        } else {
            quote! {}
        };
        let text_stmt = if text_format {
            quote! { __col = __col.as_text(); }
        } else {
            quote! {}
        };
        let image_stmt = if image {
            quote! { __col = __col.image(); }
        } else {
            quote! {}
        };

        column_defs.push(quote! {{
            let mut __col = ::excel::ColumnDef::new(#header_lit);
            let __aliases: ::std::vec::Vec<::std::string::String> =
                vec![#(#aliases.to_string()),*];
            if !__aliases.is_empty() {
                __col = __col.aliases(__aliases);
            }
            __col = __col.kind(::excel::ColumnKind::#kind_ident);
            #required_stmt
            #text_stmt
            #image_stmt
            #width_stmt
            #format_stmt
            #default_stmt
            #note_stmt
            let __dropdown: ::std::vec::Vec<::std::string::String> =
                vec![#(#dropdown.to_string()),*];
            if !__dropdown.is_empty() {
                __col = __col.dropdown(__dropdown);
            }
            __col
        }});

        // ── from_row() ──
        let value_ident =
            syn::Ident::new(&format!("__value_{index}"), proc_macro2::Span::call_site());
        let header_for_extract = header.clone();
        let parse_expr = if attrs.image {
            quote! {
                let __parsed = match row.images_of(__header).first() {
                    ::core::option::Option::Some(__img) => {
                        let __cell = ::excel::Cell {
                            row: row.row_index,
                            col: __img.column,
                            value: ::excel::CellValue::Image(__img.image()),
                            style: 0,
                        };
                        <#ty as ::excel::FromCell>::from_cell(&__cell)
                    }
                    ::core::option::Option::None => {
                        <#ty as ::excel::FromCell>::from_missing_column()
                    }
                };
            }
        } else {
            quote! {
                let __parsed = match row.cell_of(__header) {
                    ::core::option::Option::Some(__cell) => {
                        <#ty as ::excel::FromCell>::from_cell(__cell)
                    }
                    ::core::option::Option::None => {
                        <#ty as ::excel::FromCell>::from_missing_column()
                    }
                };
            }
        };
        extract_stmts.push(quote! {
            let #value_ident = {
                let __header: &str = #header_for_extract;
                #parse_expr
                match __parsed {
                    ::core::result::Result::Ok(__value) => ::core::option::Option::Some(__value),
                    ::core::result::Result::Err(__message) => {
                        let __file_header = row
                            .header
                            .index_of(__header)
                            .and_then(|__col| row.header.file_headers.get(__col as usize).cloned())
                            .unwrap_or_default();
                        __errors.push(::excel::ColumnError {
                            column: __header.to_string(),
                            file_header: __file_header,
                            message: __message,
                        });
                        ::core::option::Option::None
                    }
                }
            };
        });
        ctor_fields.push(quote! {
            #ident: #value_ident.expect("解析失败时已提前返回")
        });

        // ── to_row() ──
        to_row_exprs.push(quote! {
            ::excel::IntoCell::into_cell(&self.#ident)
        });
    }

    let dynamic_stmt = if has_dynamic {
        quote! {
            let __dynamic_values: ::std::vec::Vec<(::std::string::String, ::excel::CellValue)> = row
                .header
                .file_headers
                .iter()
                .enumerate()
                .filter(|(__i, __header)| {
                    !__header.is_empty()
                        && !row
                            .header
                            .columns
                            .iter()
                            .any(|__col| __col.file_index == ::core::option::Option::Some(*__i as u16))
                })
                .map(|(__i, __header)| {
                    (
                        __header.clone(),
                        row.cell_at(__i as u16)
                            .map(|__cell| __cell.value.clone())
                            .unwrap_or(::excel::CellValue::Empty),
                    )
                })
                .collect();
        }
    } else {
        quote! {}
    };

    let sheet_impl = match sheet_name {
        Some(sheet) => quote! {
            fn sheet_name() -> ::core::option::Option<&'static str> {
                ::core::option::Option::Some(#sheet)
            }
        },
        None => quote! {},
    };

    Ok(quote! {
        #[automatically_derived]
        impl ::excel::ExcelRow for #name {
            fn columns() -> ::std::vec::Vec<::excel::ColumnDef> {
                vec![#(#column_defs),*]
            }

            fn from_row(
                row: &::excel::RowData<'_>,
            ) -> ::core::result::Result<Self, ::std::vec::Vec<::excel::ColumnError>> {
                let mut __errors: ::std::vec::Vec<::excel::ColumnError> = ::std::vec::Vec::new();
                #dynamic_stmt
                #(#extract_stmts)*
                if !__errors.is_empty() {
                    return ::core::result::Result::Err(__errors);
                }
                ::core::result::Result::Ok(Self {
                    #(#ctor_fields),*
                })
            }

            fn to_row(&self) -> ::std::vec::Vec<::excel::CellValue> {
                vec![#(#to_row_exprs),*]
            }

            #sheet_impl
        }
    })
}

fn parse_field_attrs(field: &Field) -> syn::Result<FieldAttrs> {
    let mut out = FieldAttrs::default();
    for attr in &field.attrs {
        if !attr.path().is_ident("excel") {
            continue;
        }
        attr.parse_nested_meta(|meta| {
            if meta.path.is_ident("header") {
                out.header = Some(meta.value()?.parse::<LitStr>()?.value());
                return Ok(());
            }
            if meta.path.is_ident("alias") {
                let value = meta.value()?;
                if value.peek(syn::token::Bracket) {
                    let content;
                    syn::bracketed!(content in value);
                    let lits =
                        content.parse_terminated(|p| p.parse::<LitStr>(), Token![,])?;
                    out.aliases.extend(lits.into_iter().map(|l| l.value()));
                } else {
                    out.aliases.push(value.parse::<LitStr>()?.value());
                }
                return Ok(());
            }
            if meta.path.is_ident("dropdown") {
                let value = meta.value()?;
                let content;
                syn::bracketed!(content in value);
                let lits = content.parse_terminated(|p| p.parse::<LitStr>(), Token![,])?;
                out.dropdown.extend(lits.into_iter().map(|l| l.value()));
                return Ok(());
            }
            if meta.path.is_ident("required") {
                out.required = true;
                out.required_set = true;
                return Ok(());
            }
            if meta.path.is_ident("text") {
                out.text_format = true;
                return Ok(());
            }
            if meta.path.is_ident("image") {
                out.image = true;
                return Ok(());
            }
            if meta.path.is_ident("ignore") {
                out.ignore = true;
                return Ok(());
            }
            if meta.path.is_ident("dynamic") {
                out.dynamic = true;
                return Ok(());
            }
            if meta.path.is_ident("width") {
                let value = meta.value()?;
                let lit: Lit = value.parse()?;
                out.width = Some(match lit {
                    Lit::Float(f) => f.base10_parse::<f64>()?,
                    Lit::Int(i) => i.base10_parse::<f64>()?,
                    other => {
                        return Err(syn::Error::new_spanned(other, "width 需要数字"));
                    }
                });
                return Ok(());
            }
            for (key, slot) in [
                ("format", 0usize),
                ("default", 1),
                ("note", 2),
                ("kind", 3),
            ] {
                if meta.path.is_ident(key) {
                    let value = meta.value()?.parse::<LitStr>()?.value();
                    match slot {
                        0 => out.format = Some(value),
                        1 => out.default = Some(value),
                        2 => out.note = Some(value),
                        _ => out.kind = Some(value),
                    }
                    return Ok(());
                }
            }
            Err(meta.error(
                "不支持的 excel 属性（支持 header/alias/required/kind/width/format/text/image/ignore/dynamic/default/note/dropdown）",
            ))
        })?;
    }
    Ok(out)
}

/// 语义类型：显式 kind 优先，其次按 Rust 类型推断
fn resolve_kind(attrs: &FieldAttrs, ty: &Type) -> &'static str {
    if attrs.image {
        return "Image";
    }
    if let Some(kind) = &attrs.kind {
        let normalized = match kind.to_ascii_lowercase().as_str() {
            "text" | "string" => "Text",
            "integer" | "int" => "Integer",
            "number" | "float" | "double" => "Number",
            "decimal" | "money" => "Decimal",
            "bool" | "boolean" => "Bool",
            "date" => "Date",
            "datetime" | "date_time" => "DateTime",
            "image" => "Image",
            _ => "Any",
        };
        return normalized;
    }
    match base_type_name(ty).as_deref() {
        Some("String") => "Text",
        Some("i8") | Some("i16") | Some("i32") | Some("i64") | Some("u8") | Some("u16")
        | Some("u32") | Some("u64") | Some("isize") | Some("usize") => "Integer",
        Some("f32") | Some("f64") => "Number",
        Some("Decimal") => "Decimal",
        Some("bool") => "Bool",
        Some("NaiveDate") => "Date",
        Some("NaiveDateTime") | Some("DateTime") => "DateTime",
        Some("CellImage") | Some("RowImage") => "Image",
        _ => "Any",
    }
}

/// 取（剥掉 `Option<>` 之后的）类型名
fn base_type_name(ty: &Type) -> Option<String> {
    match ty {
        Type::Path(path) => {
            let last = path.path.segments.last()?;
            if last.ident == "Option" {
                if let syn::PathArguments::AngleBracketed(args) = &last.arguments {
                    if let Some(syn::GenericArgument::Type(inner)) = args.args.first() {
                        return base_type_name(inner);
                    }
                }
            }
            Some(last.ident.to_string())
        }
        Type::Reference(reference) => base_type_name(&reference.elem),
        _ => None,
    }
}

// ───────────────────────── 执行器注册派生宏 ─────────────────────────

/// `#[derive(ExcelExecutor)]`：把行模型注册成「执行器工厂」
///
/// 在业务模型上（通常与 `#[derive(ExcelRow)]` 组合使用）挂一个注册名即可：
///
/// ```ignore
/// use excel::{ExcelExecutor, ExcelRow};
/// use serde::Serialize;
///
/// #[derive(ExcelRow, ExcelExecutor, Serialize)]
/// #[excel(sheet = "产品导入", register = "product")]
/// struct ProductRow {
///     #[excel(header = "品名", required)]
///     name: String,
/// }
///
/// // 业务端零配置直接调用（自动创建 runner / 注册进全局表）：
/// let exec = excel::executor_for::<ProductRow>();
/// // 或按注册名字符串调度（Web 层路由用）：
/// let erased = excel::executor_by_name("product");
/// ```
///
/// 需要实现 `serde::Serialize`：注册进全局表后要能按名返回 JSON 预览。
/// 属性（挂在 `#[excel(...)]` 上，可与 [`derive(ExcelRow)`](macro@ExcelRow) 的 `sheet` 同时写）：
///
/// | 属性 | 必填 | 说明 |
/// |---|---|---|
/// | `register = "name"` | 是 | 全局注册名（按名调度的 key） |
/// | `sheet = "sheet"` | 否 | 默认工作表名（缺省退回 `ExcelRow::sheet_name`） |
#[proc_macro_derive(ExcelExecutor, attributes(excel))]
pub fn derive_excel_executor(input: TokenStream) -> TokenStream {
    let input = parse_macro_input!(input as DeriveInput);
    match expand_executor(&input) {
        Ok(tokens) => tokens.into(),
        Err(err) => err.to_compile_error().into(),
    }
}

fn expand_executor(input: &DeriveInput) -> syn::Result<proc_macro2::TokenStream> {
    if !matches!(input.data, Data::Struct(_)) {
        return Err(syn::Error::new_spanned(
            input,
            "ExcelExecutor 只能派生在结构体上",
        ));
    }
    let ident = &input.ident;

    let mut register: Option<LitStr> = None;
    let mut _sheet: Option<String> = None;
    for attr in &input.attrs {
        if !attr.path().is_ident("excel") {
            continue;
        }
        attr.parse_nested_meta(|meta| {
            if meta.path.is_ident("register") {
                register = Some(meta.value()?.parse::<LitStr>()?);
                Ok(())
            } else if meta.path.is_ident("sheet") {
                _sheet = Some(meta.value()?.parse::<LitStr>()?.value());
                Ok(())
            } else {
                Err(meta.error("ExcelExecutor 只支持 #[excel(register = \"...\", sheet = \"...\")]"))
            }
        })?;
    }
    let register = register.ok_or_else(|| {
        syn::Error::new_spanned(
            ident,
            "ExcelExecutor 需要 #[excel(register = \"注册名\")]，例如 register = \"product\"",
        )
    })?;

    // 全局唯一命名：同一模块里不可能出现两个同名类型
    let make_fn = syn::Ident::new(
        &format!("__excel_executor_make_{ident}"),
        proc_macro2::Span::call_site(),
    );

    Ok(quote! {
        #[automatically_derived]
        impl ::excel::ExcelExecutor for #ident {
            const EXECUTOR_NAME: &'static str = #register;
        }

        #[automatically_derived]
        impl #ident {
            /// 本模型的通用执行器：预览 / 导入 / 导出 / 模板一条龙，业务端无需手拼 runner
            pub fn executor() -> ::excel::Executor<Self> {
                ::excel::executor_for::<Self>()
            }
        }

        #[doc(hidden)]
        #[allow(non_snake_case)]
        fn #make_fn() -> ::std::boxed::Box<dyn ::excel::ErasedExecutor> {
            ::std::boxed::Box::new(::excel::executor_for::<#ident>())
        }

        #[allow(non_upper_case_globals, unused)]
        const _: () = {
            ::excel::inventory::submit! {
                ::excel::ExecutorEntry {
                    name: #register,
                    type_name: ::core::stringify!(#ident),
                    make: #make_fn,
                }
            }
        };
    })
}
