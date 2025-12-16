//! High-performance XLSX XML parser using WebAssembly
//!
//! This module provides fast XML parsing for XLSX files, replacing the
//! regex-based JavaScript parser with a streaming XML parser.

use quick_xml::events::Event;
use quick_xml::Reader;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use wasm_bindgen::prelude::*;

#[cfg(feature = "console_error_panic_hook")]
pub use console_error_panic_hook::set_once as set_panic_hook;

/// Initialize the WASM module (call once at startup)
#[wasm_bindgen]
pub fn init() {
    #[cfg(feature = "console_error_panic_hook")]
    console_error_panic_hook::set_once();
}

/// Parsed cell data from worksheet XML
#[derive(Debug, Serialize, Deserialize)]
pub struct ParsedCell {
    pub reference: String,
    pub cell_type: Option<String>,
    pub style_index: Option<u32>,
    pub value: Option<String>,
    pub formula: Option<String>,
}

/// Parsed row data
#[derive(Debug, Serialize, Deserialize)]
pub struct ParsedRow {
    pub row_num: u32,
    pub cells: Vec<ParsedCell>,
    pub height: Option<f64>,
    pub hidden: bool,
}

/// Parsed worksheet data
#[derive(Debug, Serialize, Deserialize)]
pub struct ParsedWorksheet {
    pub rows: Vec<ParsedRow>,
    pub merge_cells: Vec<String>,
    pub hyperlinks: Vec<ParsedHyperlink>,
    pub col_widths: HashMap<u32, f64>,
}

/// Parsed hyperlink
#[derive(Debug, Serialize, Deserialize)]
pub struct ParsedHyperlink {
    pub reference: String,
    pub rid: Option<String>,
    pub location: Option<String>,
    pub display: Option<String>,
    pub tooltip: Option<String>,
}

/// Parse worksheet XML and return structured data
#[wasm_bindgen]
pub fn parse_worksheet(xml: &str) -> JsValue {
    let result = parse_worksheet_impl(xml);
    serde_wasm_bindgen::to_value(&result).unwrap_or(JsValue::NULL)
}

fn parse_worksheet_impl(xml: &str) -> ParsedWorksheet {
    let mut reader = Reader::from_str(xml);
    reader.trim_text(true);

    let mut worksheet = ParsedWorksheet {
        rows: Vec::new(),
        merge_cells: Vec::new(),
        hyperlinks: Vec::new(),
        col_widths: HashMap::new(),
    };

    let mut buf = Vec::new();
    let mut current_row: Option<ParsedRow> = None;
    let mut current_cell: Option<ParsedCell> = None;
    let mut in_value = false;
    let mut in_formula = false;
    let mut in_inline_str = false;
    let mut text_content = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.local_name().as_ref() {
                    b"row" => {
                        let mut row = ParsedRow {
                            row_num: 0,
                            cells: Vec::new(),
                            height: None,
                            hidden: false,
                        };

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        row.row_num = val.parse().unwrap_or(0);
                                    }
                                }
                                b"ht" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        row.height = val.parse().ok();
                                    }
                                }
                                b"hidden" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        row.hidden = val == "1" || val == "true";
                                    }
                                }
                                _ => {}
                            }
                        }

                        current_row = Some(row);
                    }
                    b"c" => {
                        let mut cell = ParsedCell {
                            reference: String::new(),
                            cell_type: None,
                            style_index: None,
                            value: None,
                            formula: None,
                        };

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"r" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        cell.reference = val.to_string();
                                    }
                                }
                                b"t" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        cell.cell_type = Some(val.to_string());
                                    }
                                }
                                b"s" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        cell.style_index = val.parse().ok();
                                    }
                                }
                                _ => {}
                            }
                        }

                        current_cell = Some(cell);
                    }
                    b"v" => {
                        in_value = true;
                        text_content.clear();
                    }
                    b"f" => {
                        in_formula = true;
                        text_content.clear();
                    }
                    b"is" => {
                        in_inline_str = true;
                        text_content.clear();
                    }
                    b"t" if in_inline_str => {
                        // Text within inline string - handled by Text event
                    }
                    b"col" => {
                        let mut min: Option<u32> = None;
                        let mut max: Option<u32> = None;
                        let mut width: Option<f64> = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"min" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        min = val.parse().ok();
                                    }
                                }
                                b"max" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        max = val.parse().ok();
                                    }
                                }
                                b"width" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        width = val.parse().ok();
                                    }
                                }
                                _ => {}
                            }
                        }

                        if let (Some(min_col), Some(max_col), Some(w)) = (min, max, width) {
                            for col in min_col..=max_col {
                                worksheet.col_widths.insert(col, w);
                            }
                        }
                    }
                    b"mergeCell" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"ref" {
                                if let Ok(val) = std::str::from_utf8(&attr.value) {
                                    worksheet.merge_cells.push(val.to_string());
                                }
                            }
                        }
                    }
                    b"hyperlink" => {
                        let mut hyperlink = ParsedHyperlink {
                            reference: String::new(),
                            rid: None,
                            location: None,
                            display: None,
                            tooltip: None,
                        };

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"ref" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        hyperlink.reference = val.to_string();
                                    }
                                }
                                b"location" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        hyperlink.location = Some(val.to_string());
                                    }
                                }
                                b"display" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        hyperlink.display = Some(val.to_string());
                                    }
                                }
                                b"tooltip" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        hyperlink.tooltip = Some(val.to_string());
                                    }
                                }
                                _ => {
                                    // Check for r:id in namespace-prefixed attributes
                                    if let Ok(key) = std::str::from_utf8(attr.key.as_ref()) {
                                        if key.ends_with(":id") || key == "id" {
                                            if let Ok(val) = std::str::from_utf8(&attr.value) {
                                                hyperlink.rid = Some(val.to_string());
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if !hyperlink.reference.is_empty() {
                            worksheet.hyperlinks.push(hyperlink);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"row" => {
                    if let Some(row) = current_row.take() {
                        worksheet.rows.push(row);
                    }
                }
                b"c" => {
                    if let Some(cell) = current_cell.take() {
                        if let Some(ref mut row) = current_row {
                            row.cells.push(cell);
                        }
                    }
                }
                b"v" => {
                    in_value = false;
                    if let Some(ref mut cell) = current_cell {
                        cell.value = Some(text_content.clone());
                    }
                }
                b"f" => {
                    in_formula = false;
                    if let Some(ref mut cell) = current_cell {
                        if !text_content.is_empty() {
                            cell.formula = Some(text_content.clone());
                        }
                    }
                }
                b"is" => {
                    in_inline_str = false;
                    if let Some(ref mut cell) = current_cell {
                        cell.value = Some(text_content.clone());
                    }
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_value || in_formula || in_inline_str {
                    if let Ok(text) = e.unescape() {
                        text_content.push_str(&text);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    worksheet
}

/// Parse shared strings XML
#[wasm_bindgen]
pub fn parse_shared_strings(xml: &str) -> JsValue {
    let result = parse_shared_strings_impl(xml);
    serde_wasm_bindgen::to_value(&result).unwrap_or(JsValue::NULL)
}

fn parse_shared_strings_impl(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    reader.trim_text(false); // Preserve whitespace in strings

    let mut strings: Vec<String> = Vec::new();
    let mut buf = Vec::new();
    let mut in_si = false;
    let mut in_t = false;
    let mut current_string = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"si" => {
                    in_si = true;
                    current_string.clear();
                }
                b"t" if in_si => {
                    in_t = true;
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"si" => {
                    in_si = false;
                    strings.push(current_string.clone());
                }
                b"t" => {
                    in_t = false;
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_t {
                    if let Ok(text) = e.unescape() {
                        current_string.push_str(&text);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    strings
}

/// Style definition from styles.xml
#[derive(Debug, Serialize, Deserialize, Default)]
pub struct ParsedStyle {
    pub num_fmt_id: Option<u32>,
    pub font_id: Option<u32>,
    pub fill_id: Option<u32>,
    pub border_id: Option<u32>,
    pub xf_id: Option<u32>,
    pub apply_number_format: bool,
    pub apply_font: bool,
    pub apply_fill: bool,
    pub apply_border: bool,
    pub apply_alignment: bool,
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: bool,
    pub text_rotation: Option<i32>,
    pub indent: Option<u32>,
}

/// Font definition
#[derive(Debug, Serialize, Deserialize, Default)]
pub struct ParsedFont {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub size: Option<f64>,
    pub color: Option<String>,
    pub name: Option<String>,
}

/// Fill definition
#[derive(Debug, Serialize, Deserialize, Default)]
pub struct ParsedFill {
    pub pattern_type: Option<String>,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
}

/// Border definition
#[derive(Debug, Serialize, Deserialize, Default)]
pub struct ParsedBorder {
    pub left_style: Option<String>,
    pub left_color: Option<String>,
    pub right_style: Option<String>,
    pub right_color: Option<String>,
    pub top_style: Option<String>,
    pub top_color: Option<String>,
    pub bottom_style: Option<String>,
    pub bottom_color: Option<String>,
}

/// Parsed styles data
#[derive(Debug, Serialize, Deserialize, Default)]
pub struct ParsedStyles {
    pub cell_xfs: Vec<ParsedStyle>,
    pub fonts: Vec<ParsedFont>,
    pub fills: Vec<ParsedFill>,
    pub borders: Vec<ParsedBorder>,
    pub num_fmts: HashMap<u32, String>,
}

/// Parse styles.xml
#[wasm_bindgen]
pub fn parse_styles(xml: &str) -> JsValue {
    let result = parse_styles_impl(xml);
    serde_wasm_bindgen::to_value(&result).unwrap_or(JsValue::NULL)
}

fn parse_styles_impl(xml: &str) -> ParsedStyles {
    let mut reader = Reader::from_str(xml);
    reader.trim_text(true);

    let mut styles = ParsedStyles::default();
    let mut buf = Vec::new();

    let mut in_cell_xfs = false;
    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_borders = false;
    let mut in_num_fmts = false;

    let mut current_font: Option<ParsedFont> = None;
    let mut current_fill: Option<ParsedFill> = None;
    let mut current_border: Option<ParsedBorder> = None;
    let mut in_pattern_fill = false;
    let mut current_border_side: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                match e.local_name().as_ref() {
                    b"cellXfs" => in_cell_xfs = true,
                    b"fonts" => in_fonts = true,
                    b"fills" => in_fills = true,
                    b"borders" => in_borders = true,
                    b"numFmts" => in_num_fmts = true,
                    b"xf" if in_cell_xfs => {
                        let mut style = ParsedStyle::default();

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        style.num_fmt_id = val.parse().ok();
                                    }
                                }
                                b"fontId" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        style.font_id = val.parse().ok();
                                    }
                                }
                                b"fillId" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        style.fill_id = val.parse().ok();
                                    }
                                }
                                b"borderId" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        style.border_id = val.parse().ok();
                                    }
                                }
                                b"xfId" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        style.xf_id = val.parse().ok();
                                    }
                                }
                                b"applyNumberFormat" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        style.apply_number_format = val == "1" || val == "true";
                                    }
                                }
                                b"applyFont" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        style.apply_font = val == "1" || val == "true";
                                    }
                                }
                                b"applyFill" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        style.apply_fill = val == "1" || val == "true";
                                    }
                                }
                                b"applyBorder" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        style.apply_border = val == "1" || val == "true";
                                    }
                                }
                                b"applyAlignment" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        style.apply_alignment = val == "1" || val == "true";
                                    }
                                }
                                _ => {}
                            }
                        }

                        styles.cell_xfs.push(style);
                    }
                    b"alignment" if in_cell_xfs => {
                        if let Some(style) = styles.cell_xfs.last_mut() {
                            for attr in e.attributes().flatten() {
                                match attr.key.as_ref() {
                                    b"horizontal" => {
                                        if let Ok(val) = std::str::from_utf8(&attr.value) {
                                            style.horizontal = Some(val.to_string());
                                        }
                                    }
                                    b"vertical" => {
                                        if let Ok(val) = std::str::from_utf8(&attr.value) {
                                            style.vertical = Some(val.to_string());
                                        }
                                    }
                                    b"wrapText" => {
                                        if let Ok(val) = std::str::from_utf8(&attr.value) {
                                            style.wrap_text = val == "1" || val == "true";
                                        }
                                    }
                                    b"textRotation" => {
                                        if let Ok(val) = std::str::from_utf8(&attr.value) {
                                            style.text_rotation = val.parse().ok();
                                        }
                                    }
                                    b"indent" => {
                                        if let Ok(val) = std::str::from_utf8(&attr.value) {
                                            style.indent = val.parse().ok();
                                        }
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    b"font" if in_fonts => {
                        current_font = Some(ParsedFont::default());
                    }
                    b"b" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.bold = true;
                        }
                    }
                    b"i" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.italic = true;
                        }
                    }
                    b"u" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.underline = true;
                        }
                    }
                    b"strike" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            font.strikethrough = true;
                        }
                    }
                    b"sz" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        font.size = val.parse().ok();
                                    }
                                }
                            }
                        }
                    }
                    b"color" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"rgb" {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        font.color = Some(val.to_string());
                                    }
                                }
                            }
                        }
                    }
                    b"name" if current_font.is_some() => {
                        if let Some(ref mut font) = current_font {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"val" {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        font.name = Some(val.to_string());
                                    }
                                }
                            }
                        }
                    }
                    b"fill" if in_fills => {
                        current_fill = Some(ParsedFill::default());
                    }
                    b"patternFill" if current_fill.is_some() => {
                        in_pattern_fill = true;
                        if let Some(ref mut fill) = current_fill {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"patternType" {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        fill.pattern_type = Some(val.to_string());
                                    }
                                }
                            }
                        }
                    }
                    b"fgColor" if in_pattern_fill => {
                        if let Some(ref mut fill) = current_fill {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"rgb" {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        fill.fg_color = Some(val.to_string());
                                    }
                                }
                            }
                        }
                    }
                    b"bgColor" if in_pattern_fill => {
                        if let Some(ref mut fill) = current_fill {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"rgb" {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        fill.bg_color = Some(val.to_string());
                                    }
                                }
                            }
                        }
                    }
                    b"border" if in_borders => {
                        current_border = Some(ParsedBorder::default());
                    }
                    b"left" | b"right" | b"top" | b"bottom" if current_border.is_some() => {
                        let side = std::str::from_utf8(e.local_name().as_ref())
                            .unwrap_or("")
                            .to_string();
                        current_border_side = Some(side.clone());

                        if let Some(ref mut border) = current_border {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"style" {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        match side.as_str() {
                                            "left" => border.left_style = Some(val.to_string()),
                                            "right" => border.right_style = Some(val.to_string()),
                                            "top" => border.top_style = Some(val.to_string()),
                                            "bottom" => border.bottom_style = Some(val.to_string()),
                                            _ => {}
                                        }
                                    }
                                }
                            }
                        }
                    }
                    b"color" if current_border_side.is_some() => {
                        if let Some(ref mut border) = current_border {
                            for attr in e.attributes().flatten() {
                                if attr.key.as_ref() == b"rgb" {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        match current_border_side.as_deref() {
                                            Some("left") => {
                                                border.left_color = Some(val.to_string())
                                            }
                                            Some("right") => {
                                                border.right_color = Some(val.to_string())
                                            }
                                            Some("top") => border.top_color = Some(val.to_string()),
                                            Some("bottom") => {
                                                border.bottom_color = Some(val.to_string())
                                            }
                                            _ => {}
                                        }
                                    }
                                }
                            }
                        }
                    }
                    b"numFmt" if in_num_fmts => {
                        let mut id: Option<u32> = None;
                        let mut code: Option<String> = None;

                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        id = val.parse().ok();
                                    }
                                }
                                b"formatCode" => {
                                    if let Ok(val) = std::str::from_utf8(&attr.value) {
                                        code = Some(val.to_string());
                                    }
                                }
                                _ => {}
                            }
                        }

                        if let (Some(id), Some(code)) = (id, code) {
                            styles.num_fmts.insert(id, code);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"cellXfs" => in_cell_xfs = false,
                b"fonts" => in_fonts = false,
                b"fills" => in_fills = false,
                b"borders" => in_borders = false,
                b"numFmts" => in_num_fmts = false,
                b"font" if in_fonts => {
                    if let Some(font) = current_font.take() {
                        styles.fonts.push(font);
                    }
                }
                b"fill" if in_fills => {
                    if let Some(fill) = current_fill.take() {
                        styles.fills.push(fill);
                    }
                    in_pattern_fill = false;
                }
                b"patternFill" => {
                    in_pattern_fill = false;
                }
                b"border" if in_borders => {
                    if let Some(border) = current_border.take() {
                        styles.borders.push(border);
                    }
                }
                b"left" | b"right" | b"top" | b"bottom" => {
                    current_border_side = None;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    styles
}

/// Workbook sheet info
#[derive(Debug, Serialize, Deserialize)]
pub struct ParsedSheetInfo {
    pub name: String,
    pub sheet_id: u32,
    pub rid: String,
    pub state: Option<String>,
}

/// Parse workbook.xml to get sheet list
#[wasm_bindgen]
pub fn parse_workbook(xml: &str) -> JsValue {
    let result = parse_workbook_impl(xml);
    serde_wasm_bindgen::to_value(&result).unwrap_or(JsValue::NULL)
}

fn parse_workbook_impl(xml: &str) -> Vec<ParsedSheetInfo> {
    let mut reader = Reader::from_str(xml);
    reader.trim_text(true);

    let mut sheets: Vec<ParsedSheetInfo> = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"sheet" {
                    let mut sheet = ParsedSheetInfo {
                        name: String::new(),
                        sheet_id: 0,
                        rid: String::new(),
                        state: None,
                    };

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"name" => {
                                if let Ok(val) = std::str::from_utf8(&attr.value) {
                                    sheet.name = val.to_string();
                                }
                            }
                            b"sheetId" => {
                                if let Ok(val) = std::str::from_utf8(&attr.value) {
                                    sheet.sheet_id = val.parse().unwrap_or(0);
                                }
                            }
                            b"state" => {
                                if let Ok(val) = std::str::from_utf8(&attr.value) {
                                    sheet.state = Some(val.to_string());
                                }
                            }
                            _ => {
                                // Check for r:id
                                if let Ok(key) = std::str::from_utf8(attr.key.as_ref()) {
                                    if key.ends_with(":id") || key == "id" {
                                        if let Ok(val) = std::str::from_utf8(&attr.value) {
                                            sheet.rid = val.to_string();
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if !sheet.name.is_empty() {
                        sheets.push(sheet);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    sheets
}

/// Relationship info
#[derive(Debug, Serialize, Deserialize)]
pub struct ParsedRelationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

/// Parse relationships file (.rels)
#[wasm_bindgen]
pub fn parse_relationships(xml: &str) -> JsValue {
    let result = parse_relationships_impl(xml);
    serde_wasm_bindgen::to_value(&result).unwrap_or(JsValue::NULL)
}

fn parse_relationships_impl(xml: &str) -> Vec<ParsedRelationship> {
    let mut reader = Reader::from_str(xml);
    reader.trim_text(true);

    let mut rels: Vec<ParsedRelationship> = Vec::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let mut rel = ParsedRelationship {
                        id: String::new(),
                        rel_type: String::new(),
                        target: String::new(),
                        target_mode: None,
                    };

                    for attr in e.attributes().flatten() {
                        match attr.key.as_ref() {
                            b"Id" => {
                                if let Ok(val) = std::str::from_utf8(&attr.value) {
                                    rel.id = val.to_string();
                                }
                            }
                            b"Type" => {
                                if let Ok(val) = std::str::from_utf8(&attr.value) {
                                    rel.rel_type = val.to_string();
                                }
                            }
                            b"Target" => {
                                if let Ok(val) = std::str::from_utf8(&attr.value) {
                                    rel.target = val.to_string();
                                }
                            }
                            b"TargetMode" => {
                                if let Ok(val) = std::str::from_utf8(&attr.value) {
                                    rel.target_mode = Some(val.to_string());
                                }
                            }
                            _ => {}
                        }
                    }

                    if !rel.id.is_empty() {
                        rels.push(rel);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    rels
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_shared_strings() {
        let xml = r#"<?xml version="1.0"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>Hello</t></si>
            <si><t>World</t></si>
            <si><r><t>Rich</t></r><r><t>Text</t></r></si>
        </sst>"#;

        let strings = parse_shared_strings_impl(xml);
        assert_eq!(strings.len(), 3);
        assert_eq!(strings[0], "Hello");
        assert_eq!(strings[1], "World");
        assert_eq!(strings[2], "RichText");
    }

    #[test]
    fn test_parse_worksheet() {
        let xml = r#"<?xml version="1.0"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1" t="s"><v>0</v></c>
                    <c r="B1"><v>42</v></c>
                </row>
            </sheetData>
        </worksheet>"#;

        let worksheet = parse_worksheet_impl(xml);
        assert_eq!(worksheet.rows.len(), 1);
        assert_eq!(worksheet.rows[0].cells.len(), 2);
        assert_eq!(worksheet.rows[0].cells[0].reference, "A1");
        assert_eq!(worksheet.rows[0].cells[0].cell_type, Some("s".to_string()));
        assert_eq!(worksheet.rows[0].cells[0].value, Some("0".to_string()));
    }

    #[test]
    fn test_parse_workbook() {
        let xml = r#"<?xml version="1.0"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheets>
                <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
            </sheets>
        </workbook>"#;

        let sheets = parse_workbook_impl(xml);
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].name, "Sheet1");
        assert_eq!(sheets[1].name, "Sheet2");
    }
}
