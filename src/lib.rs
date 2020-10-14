use wasm_bindgen;
use wasm_bindgen::prelude::*;

use serde_derive::{Serialize};

use quick_xml;
use zip;

use quick_xml::events::attributes::{Attribute};
use quick_xml::events::{Event};
use quick_xml::Reader as XmlReader;
use zip::read::{ZipArchive, ZipFile};

use std::io::Cursor;
use std::io::BufReader;

use std::collections::HashMap;

pub mod border;
use crate::border::{Border, BorderPosition};
pub mod range;
use crate::range::{Range, cell_index_to_offsets};

type XlsReader<'a> = XmlReader<BufReader<ZipFile<'a>>>;
type Sheet = (String, String);
type Dict = HashMap<String, String>;

const DEFAULT_CELL_WIDTH: f32 = 15.75;
const DEFAULT_CELL_HEIGHT: f32 = 14.25;
const WIDTH_COEF: f32 = 8.5;
const HEIGHT_COEF: f32 = 0.75;
const PT_COEF: f32 = 0.75;

#[derive(PartialEq)]
enum StyleXMLPath {
    Any,
    Font,
    Fill,
    Border,
    CellXfs,
    Xf,
}

pub const WITH_FORMULAS: u32   = 1;

#[derive(PartialEq)]
enum SharedStringXMLPath {
    Any,
    Si,
}

#[derive(Debug)]
enum XlsxError {
    Default,
    FileNotFound(String)
}


#[derive(Serialize)]
pub struct ColumnData {
    pub width: u32,
}

#[derive(Serialize)]
pub struct RowData {
    pub height: u32,
}

#[derive(Serialize)]
pub struct CellCoords {
    pub column: u32,
    pub row: u32,
}

#[derive(Serialize)]
pub struct MergedCell {
    pub from: CellCoords,
    pub to: CellCoords,
}

#[derive(Serialize)]
pub struct Cell {
    pub v: Option<String>,
    pub s: u32,
}

impl Cell {
    pub fn new() -> Cell {
        Cell {
            v: None,
            s: 0,
        }
    }
}

#[derive(Serialize)]
pub struct SheetData {
    pub name: String,
    pub cols: Vec<ColumnData>,
    pub rows: Vec<RowData>,
    pub cells: Vec<Vec<Option<Cell>>>,
    pub merged: Vec<MergedCell>,
}

impl SheetData {
    pub fn new(name: String) -> SheetData {
        SheetData {
            name,
            cols: vec!(),
            rows: vec!(),
            cells: vec!(),
            merged: vec!(),
        }
    }
}

struct SheetInfo {
    cols_count: u32,
    default_col_width: u32,
    default_row_height: u32,
    use_shared_string_for_next: bool,
}

impl SheetInfo {
    pub fn new() -> SheetInfo {
        SheetInfo {
            cols_count: 0,
            default_col_width: (DEFAULT_CELL_WIDTH * WIDTH_COEF).round() as u32,
            default_row_height: (DEFAULT_CELL_HEIGHT / HEIGHT_COEF).round() as u32,
            use_shared_string_for_next: false,
        }
    }
}

#[wasm_bindgen]
pub struct XLSX {
    shared_strings: Vec<String>,
    sheets: Vec<Sheet>,
    zip: ZipArchive<Cursor<Vec<u8>>>,
}

#[wasm_bindgen]
impl XLSX {
    pub fn new(data: Vec<u8>) -> XLSX {
        let buf = Cursor::new(data);
        let zip = ZipArchive::new(buf).unwrap();

        let mut xlsx = XLSX {
            shared_strings: vec!(),
            sheets: vec!(),
            zip
        };

        let rels = xlsx.read_relationships().unwrap();
        xlsx.read_workbook(&rels).unwrap();
        match xlsx.read_shared_strings() {
            Ok(_) => (),
            Err(_) => ()
        }

        xlsx
    }
    pub fn with_formulas() -> u32{
        return WITH_FORMULAS;
    }
    pub fn get_styles(&mut self) -> JsValue {
        let styles = self.read_style().unwrap();
        JsValue::from_serde(&styles).unwrap()
    }
    pub fn get_sheets(&self) -> Vec<JsValue> {
        self.sheets.clone().iter().map(|s| JsValue::from(&s.0)).collect()
    }
    pub fn get_sheet_data(&mut self, sheet_name: String, flags: u32) -> JsValue {
        let (name, path) = self.sheets.iter().find(|(name, _)| name == &sheet_name).unwrap().clone();
        let data = self.read_sheet(path, name, flags).unwrap();

        JsValue::from_serde(&data).unwrap()
    }
    fn read_shared_strings(&mut self) -> Result<(), XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/sharedStrings.xml") {
            None => return Ok(()),
            Some(x) => x?,
        };
        let mut buf = Vec::new();
        
        let mut xml_path = SharedStringXMLPath::Any;
        loop {
            buf.clear();
            match xml.read_event(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name() == b"si" => {
                    xml_path = SharedStringXMLPath::Si;
                }
                Ok(Event::Start(ref e)) if e.local_name() == b"t" && xml_path == SharedStringXMLPath::Si => {
                    let value = xml.read_text(e.name(), &mut Vec::new()).unwrap();
                    self.shared_strings.push(value);
                }
                Ok(Event::End(ref e)) if e.local_name() == b"si" => {
                    xml_path = SharedStringXMLPath::Any;
                },
                Ok(Event::End(ref e)) if e.local_name() == b"sst" => break,
                Err(_) => return Err(XlsxError::Default),
                _ => (),
            }
        }
        Ok(())
    }
    fn read_sheet(&mut self, path: String, sheet_name: String, flags: u32) -> Result<SheetData, XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, &path) {
            None => {
                return Err(XlsxError::FileNotFound(path))
            }
            Some(x) => x?,
        };
        let mut buf = Vec::new();

        let mut data = SheetData::new(sheet_name);
        let mut info = SheetInfo::new();

        let mut last_cell = Cell::new();

        loop {
            buf.clear();
            match xml.read_event(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name() == b"sheetFormatPr" => {
                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"tdefaultRowHeight", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap().parse::<f32>().unwrap() / HEIGHT_COEF;
                                info.default_row_height = value.round() as u32;
                            },
                            Attribute { key: b"defaultColWidth", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap().parse::<f32>().unwrap() * WIDTH_COEF;
                                info.default_col_width = value.round() as u32;
                            },
                            _ => ()
                        }
                    }
                },
                Ok(Event::Start(ref e)) if e.local_name() == b"col" => {
                    let mut min = 0;
                    let mut max = 0;
                    let mut width = 0;
                    let mut use_custom_width = false;

                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"width", .. } => {
                                width = (att.unescape_and_decode_value(&xml).unwrap().parse::<f32>().unwrap() * WIDTH_COEF).round() as u32;
                            },
                            Attribute { key: b"min", .. } => {
                                min = att.unescape_and_decode_value(&xml).unwrap().parse::<usize>().unwrap();
                            },
                            Attribute { key: b"max", .. } => {
                                max = att.unescape_and_decode_value(&xml).unwrap().parse::<usize>().unwrap();
                            },
                            Attribute { key: b"customWidth", .. } => {
                                use_custom_width = att.unescape_and_decode_value(&xml).unwrap() == "1";
                            },
                            _ => ()
                        }
                    }
                    if use_custom_width {
                        for i in data.cols.len()..max {
                            if i >= min-1 {
                                data.cols.push(ColumnData {width});
                            } else {
                                data.cols.push(ColumnData {width: info.default_col_width});
                            }
                        }
                    }
                },
                Ok(Event::Start(ref e)) if e.local_name() == b"row" => {
                    let mut use_custom_height = false;
                    let mut height = 0;
                    let mut index = 0;

                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"ht", .. } => {
                                height = (att.unescape_and_decode_value(&xml).unwrap().parse::<f32>().unwrap() / HEIGHT_COEF).round() as u32;
                            },
                            Attribute { key: b"customHeight", .. } => {
                                use_custom_height = att.unescape_and_decode_value(&xml).unwrap() == "1";
                            },
                            Attribute { key: b"r", .. } => {
                                index = att.unescape_and_decode_value(&xml).unwrap().parse::<usize>().unwrap();
                            }
                            _ => ()
                        }
                    }
                    for _ in data.cells.len()..index {
                        data.cells.push(vec!());
                    }
                    for _ in data.rows.len()..index-1 {
                        data.rows.push(RowData {height: info.default_row_height});
                    }
                    if use_custom_height {
                        data.rows.push(RowData {height});
                    } else {
                        data.rows.push(RowData {height: info.default_row_height});
                    }
                },
                Ok(Event::Start(ref e)) if e.local_name() == b"c" => {
                    info.use_shared_string_for_next = false;

                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"t", .. } => {
                                if att.unescape_and_decode_value(&xml).unwrap() == "s" {
                                    info.use_shared_string_for_next = true;
                                }
                            },
                            Attribute { key: b"s", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap().parse::<u32>().unwrap();
                                last_cell.s = value;
                            },
                            Attribute { key: b"r", value: v } => {
                                let cell_name = xml.decode(&v).into_owned();
                                let (col, _) = cell_index_to_offsets(cell_name.clone());

                                let cols = data.cells.last_mut().unwrap();

                                for _ in cols.len()..col as usize {
                                    cols.push(None);
                                }

                                if col + 1 > info.cols_count {
                                    info.cols_count = col + 1;
                                }
                            },
                            _ => ()
                        }
                    }
                },
                Ok(Event::End(ref e)) if e.local_name() == b"c" => {
                    data.cells.last_mut().unwrap().push(Some(last_cell));
                    last_cell = Cell::new();
                },
                Ok(Event::Start(ref e)) if e.local_name() == b"f" => {
                    if flags & WITH_FORMULAS > 0 {
                        let fline = &mut Vec::new();
                        let value = xml.read_text(e.name(), fline).unwrap();
                        last_cell.v = Some("=".to_owned() + &value);
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name() == b"v" => {
                    if last_cell.v.is_none(){
                        let value = xml.read_text(e.name(), &mut Vec::new()).unwrap();
                        if info.use_shared_string_for_next {
                            let index: usize = value.parse().unwrap();
                            last_cell.v = Some(self.shared_strings[index].to_owned());
                        } else {
                            last_cell.v = Some(value);
                        }
                    }
                },
                Ok(Event::Start(ref e)) if e.local_name() == b"mergeCell" => {
                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"ref", .. } => {
                                let raw_merged_cell = att.unescape_and_decode_value(&xml).unwrap();
                                let range = Range::new(raw_merged_cell);

                                let from = range.first;
                                let to = range.last;

                                let merged_cell = MergedCell {
                                    from: CellCoords { column: from.0, row: from.1 },
                                    to: CellCoords { column: to.0, row: to.1 }
                                };
                                data.merged.push(merged_cell);
                            },
                            _ => ()
                        }
                    }
                },
                Ok(Event::End(ref e)) if e.local_name() == b"worksheet" => {
                    loop {
                        match data.cells.last() {
                            Some(last) => {
                                if last.len() == 0 {
                                    data.cells.pop();
                                } else {
                                    break;
                                }
                            },
                            None => break
                        }
                    }
                    let rows_count = data.cells.len();
                    for row in data.cells.iter_mut() {
                        let missed_cols_count = info.cols_count as usize - row.len();
                        if missed_cols_count != 0 {
                            row.extend((0..missed_cols_count).map(|_| None));
                        }
                    }
                    let missed_row_data_count = rows_count as i32 - data.rows.len() as i32;
                    if missed_row_data_count < 0 {
                        data.rows = data.rows.into_iter().take(rows_count).collect();
                    } else {
                        data.rows.extend((0..missed_row_data_count).map(|_| RowData {height: info.default_row_height}));
                    }
                    let missed_col_data_count = info.cols_count as i32 - data.cols.len() as i32;
                    if missed_col_data_count < 0 {
                        data.cols = data.cols.into_iter().take(info.cols_count as usize).collect();
                    } else {
                        data.cols.extend((0..missed_col_data_count).map(|_| ColumnData {width: info.default_col_width}));
                    }
                    return Ok(data);
                },
                Err(_) => return Err(XlsxError::Default),
                _ => ()
            }
        }
    }

    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>, XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/_rels/workbook.xml.rels") {
            None => {
                return Err(XlsxError::FileNotFound(
                    String::from("xl/_rels/workbook.xml.rels"),
                ))
            },
            Some(x) => x?,
        };
        let mut relationships = HashMap::new();
        let mut buf = Vec::new();
        loop {
            buf.clear();
            match xml.read_event(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name() == b"Relationship" => {
                    let mut id = Vec::new();
                    let mut target = String::new();
                    for a in e.attributes() {
                        match a.unwrap() {
                            Attribute { key: b"Id", value: v } => {
                                id.extend_from_slice(&v);
                            },
                            Attribute { key: b"Target", value: v } => {
                                target = xml.decode(&v).into_owned();
                            },
                            _ => (),
                        }
                    }
                    relationships.insert(id, target);
                }
                Ok(Event::End(ref e)) if e.local_name() == b"Relationships" => break,
                Err(_) => return Err(XlsxError::Default),
                _ => (),
            }
        }
        Ok(relationships)
    }

    fn read_workbook(&mut self, relationships: &HashMap<Vec<u8>, String>) -> Result<(), XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/workbook.xml") {
            None => {
                return Err(XlsxError::FileNotFound(
                    String::from("xl/workbook.xml")
                ))
            },
            Some(x) => x?,
        };
        let mut buf = Vec::new();
        loop {
            buf.clear();
            match xml.read_event(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name() == b"sheet" => {
                    let mut name = String::new();
                    let mut path = String::new();
                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"name", .. } => {
                                name = att.unescape_and_decode_value(&xml).unwrap();
                            },
                            Attribute {
                                key: b"r:id",
                                value: v,
                            } => {
                                let r = &relationships[&*v][..];
                                path = if r.starts_with("/xl/") {
                                    r[1..].to_string()
                                } else if r.starts_with("xl/") {
                                    r.to_string()
                                } else {
                                    format!("xl/{}", r)
                                };
                            }
                            _ => ()
                        }
                    }
                    self.sheets.push((name, path));
                },
                Ok(Event::End(ref e)) if e.local_name() == b"workbook" => break,
                Err(_) => return Err(XlsxError::Default),
                _ => (),
            }
        }
        Ok(())
    }

    fn read_style(&mut self) -> Result<Vec<Dict>, XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/styles.xml") {
            None => {
                return Err(XlsxError::FileNotFound(
                    String::from("xl/styles.xml")
                ))
            },
            Some(x) => x?,
        };
        let mut buf = Vec::new();

        let mut xml_path: StyleXMLPath = StyleXMLPath::Any;
        let mut xml_parent_path: StyleXMLPath = StyleXMLPath::Any;

        let mut fonts: Vec<Dict> = vec!();
        let mut fills: Vec<Dict> = vec!();
        let mut borders: Vec<Dict> = vec!();
        let mut border_structs: Vec<Border> = vec!();

        let mut extra_formats: Dict = HashMap::new();

        let mut styles: Vec<Dict> = vec!();

        loop {
            buf.clear();
            match xml.read_event(&mut buf) {
                Ok(Event::Start(ref e)) if xml_parent_path == StyleXMLPath::CellXfs && e.local_name() == b"xf"  => {
                    xml_path = StyleXMLPath::Xf;
                    let mut xf = HashMap::new();
                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"fontId", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                let index: usize = value.parse().unwrap();
                                let font_style: Dict = fonts[index].clone();

                                xf.extend(font_style);
                            },
                            Attribute { key: b"borderId", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                let index: usize = value.parse().unwrap();
                                let border_style: Dict = borders[index].clone();

                                xf.extend(border_style);
                            },
                            Attribute { key: b"fillId", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                let index: usize = value.parse().unwrap();
                                let fill_style: Dict = fills[index].clone();

                                xf.extend(fill_style);
                            },
                            Attribute { key: b"numFmtId", .. } => {
                                let format_id = att.unescape_and_decode_value(&xml).unwrap();
                                let format = match get_format(&format_id) {
                                    Some(v) => v,
                                    None => extra_formats.get(&format_id).unwrap().to_owned()
                                };
                                xf.insert(String::from("format"), format);
                            },
                            _ => ()
                        }
                    }
                    styles.push(xf);
                },
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Xf && xml_parent_path == StyleXMLPath::CellXfs && e.local_name() == b"alignment" => {
                    let xf = styles.last_mut().unwrap();
                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"vertical", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                xf.insert(String::from("verticalAlign"), value);
                            },
                            Attribute { key: b"horizontal", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                xf.insert(String::from("textAlign"), value);
                            },
                            _ => ()
                        }
                    }
                },
                Ok(Event::Start(ref e)) if e.local_name() == b"numFmt" => {
                    let mut format_code = String::from("");
                    let mut format_id = String::from("");
                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"formatCode", .. } => {
                                format_code = att.unescape_and_decode_value(&xml).unwrap();
                            },
                            Attribute { key: b"numFmtId", .. } => {
                                format_id = att.unescape_and_decode_value(&xml).unwrap();
                            },
                            _ => ()
                        }
                    }
                    extra_formats.insert(format_id, format_code);
                },
                Ok(Event::Start(ref e)) if e.local_name() == b"font" => {
                    xml_path = StyleXMLPath::Font;
                    fonts.push(HashMap::new());
                },
                Ok(Event::Start(ref e)) if e.local_name() == b"fill" => {
                    xml_path = StyleXMLPath::Fill;
                    fills.push(HashMap::new());
                },
                Ok(Event::Start(ref e)) if e.local_name() == b"border" => {
                    xml_path = StyleXMLPath::Border;
                },
                Ok(Event::Start(ref e)) if e.local_name() == b"cellXfs" => {
                    xml_parent_path = StyleXMLPath::CellXfs;
                },
                // font styles
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Font && e.local_name() == b"sz"  => {
                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"val", .. } => {
                                let font = fonts.last_mut().unwrap();
                                let value = att.unescape_and_decode_value(&xml).unwrap().parse::<f32>().unwrap();
                                font.insert(String::from("fontSize"), (value / PT_COEF).round().to_string() + "px");
                            },
                            _ => ()
                        }
                    }
                },
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Font && e.local_name() ==  b"name" => {
                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"val", .. } => {
                                let font = fonts.last_mut().unwrap();
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                font.insert(String::from("fontFamily"), value);
                            },
                            _ => ()
                        }
                    } 
                },
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Font && e.local_name() == b"color" => {
                    let font = fonts.last_mut().unwrap();
                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"rgb", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                font.insert(String::from("color"), get_xlsx_rgb(value));
                            },
                            Attribute { key: b"indexed", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                font.insert(String::from("color"), get_indexed_color(&value));
                            },
                            _ => ()
                        }
                    }
                },
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Font && e.local_name() == b"b"  => {
                    let font = fonts.last_mut().unwrap();
                    font.insert(String::from("fontWeight"), String::from("bold"));
                },
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Font && e.local_name() == b"i" => {
                    let font = fonts.last_mut().unwrap();
                    font.insert(String::from("fontStyle"), String::from("italic"));
                },
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Font && e.local_name() == b"u" => {
                    let font = fonts.last_mut().unwrap();
                    if font.contains_key("textDecoration") {
                        font.insert(String::from("textDecoration"), String::from("line-through underline"));
                    } else {
                        font.insert(String::from("textDecoration"), String::from("underline"));
                    }
                },
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Font && e.local_name() == b"strike"  => {
                    let font = fonts.last_mut().unwrap();
                    if font.contains_key("textDecoration") {
                        font.insert(String::from("textDecoration"), String::from("line-through underline"));
                    } else {
                        font.insert(String::from("textDecoration"), String::from("line-through"));
                    }
                },
                // borders styles
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Border && e.local_name() == b"left" => {
                    let mut border = Border::new(BorderPosition::Left);

                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"style", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                &border.set_style(value);
                            },
                            _ => ()
                        }
                    }
                    border_structs.push(border);
                },
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Border && e.local_name() == b"right" => {
                    let mut border = Border::new(BorderPosition::Right);

                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"style", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                &border.set_style(value);
                            },
                            _ => ()
                        }
                    }
                    border_structs.push(border);
                },
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Border && e.local_name() == b"bottom" => {
                    let mut border = Border::new(BorderPosition::Bottom);

                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"style", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                &border.set_style(value);
                            },
                            _ => ()
                        }
                    }
                    border_structs.push(border);
                },
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Border && e.local_name() == b"top" => {
                    let mut border = Border::new(BorderPosition::Top);

                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"style", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                &border.set_style(value);
                            },
                            _ => ()
                        }
                    }
                    border_structs.push(border);
                },
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Border && e.local_name() == b"color" => {
                    let border = border_structs.last_mut().unwrap();
                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"rgb", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                border.set_color(get_xlsx_rgb(value));
                            },
                            Attribute { key: b"indexed", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                border.set_color(get_indexed_color(&value));
                            },
                            _ => ()
                        }
                    }
                },
                // fills
                Ok(Event::Start(ref e)) if xml_path == StyleXMLPath::Fill && e.local_name() == b"fgColor" => {
                    let fill = fills.last_mut().unwrap();
                    for a in e.attributes() {
                        let att = a.unwrap();
                        match att {
                            Attribute { key: b"rgb", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                fill.insert(String::from("background"), get_xlsx_rgb(value));
                            },
                            Attribute { key: b"indexed", .. } => {
                                let value = att.unescape_and_decode_value(&xml).unwrap();
                                fill.insert(String::from("background"), get_indexed_color(&value));
                            },
                            _ => ()
                        }
                    }
                },
                Ok(Event::End(ref e)) if e.local_name() == b"cellXfs" => {
                    xml_parent_path = StyleXMLPath::Any;
                },
                Ok(Event::End(ref e)) if xml_parent_path == StyleXMLPath::CellXfs && e.local_name() == b"xf" => {
                    xml_path = StyleXMLPath::Any;
                },
                Ok(Event::End(ref e)) if e.local_name() == b"font" => {
                    xml_path = StyleXMLPath::Any;
                },
                Ok(Event::End(ref e)) if e.local_name() == b"fill" => {
                    xml_path = StyleXMLPath::Any;
                },
                Ok(Event::End(ref e)) if e.local_name() == b"border" => {
                    xml_path = StyleXMLPath::Any;
                    let mut border = HashMap::new();

                    while border_structs.len() > 0 {
                        let border_struct = border_structs.pop().unwrap();
                        let (key, value) = border_struct.get_computed_style();
                        border.insert(key, value);
                    }
                    borders.push(border);
                },
                Ok(Event::End(ref e)) if e.local_name() == b"styleSheet" => break,
                Err(_) => return Err(XlsxError::Default),
                _ => (),
            }
        }
        Ok(styles)
    }
}

fn xml_reader<'a>(zip: &'a mut ZipArchive<Cursor<Vec<u8>>>, path: &str) -> Option<Result<XlsReader<'a>, XlsxError>> {
    match zip.by_name(path) {
        Ok(f) => {
            let mut r = XmlReader::from_reader(BufReader::new(f));
            r.check_end_names(false)
                .trim_text(false)
                .check_comments(false)
                .expand_empty_elements(true);
            Some(Ok(r))
        }
        Err(_) => None
    }
}

fn get_xlsx_rgb(argb: String) -> String {
    let raw_a = u8::from_str_radix(&argb[..2], 16).unwrap();
    let a = (raw_a as f32 / 255f32).to_string();
    let r = u8::from_str_radix(&argb[2..4], 16).unwrap().to_string();
    let g = u8::from_str_radix(&argb[4..6], 16).unwrap().to_string();
    let b = u8::from_str_radix(&argb[6..8], 16).unwrap().to_string();

    format!("rgba({},{},{},{})", r, g, b, a) 
}

fn get_indexed_color(s: &str) -> String {
    match s {
        "0" => String::from("#000000"),
        "1" => String::from("#FFFFFF"),
        "2" => String::from("#FF0000"),
        "3" => String::from("#00FF00"),
        "4" => String::from("#0000FF"),
        "5" => String::from("#FFFF00"),
        "6" => String::from("#FF00FF"),
        "7" => String::from("#00FFFF"),
        "8" => String::from("#000000"),
        "9" => String::from("#FFFFFF"),
        "10" => String::from("#FF0000"),
        "11" => String::from("#00FF00"),
        "12" => String::from("#0000FF"),
        "13" => String::from("#FFFF00"),
        "14" => String::from("#FF00FF"),
        "15" => String::from("#00FFFF"),
        "16" => String::from("#800000"),
        "17" => String::from("#008000"),
        "18" => String::from("#000080"),
        "19" => String::from("#808000"),
        "20" => String::from("#800080"),
        "21" => String::from("#008080"),
        "22" => String::from("#C0C0C0"),
        "23" => String::from("#808080"),
        "24" => String::from("#9999FF"),
        "25" => String::from("#993366"),
        "26" => String::from("#FFFFCC"),
        "27" => String::from("#CCFFFF"),
        "28" => String::from("#660066"),
        "29" => String::from("#FF8080"),
        "30" => String::from("#0066CC"),
        "31" => String::from("#CCCCFF"),
        "32" => String::from("#000080"),
        "33" => String::from("#FF00FF"),
        "34" => String::from("#FFFF00"),
        "35" => String::from("#00FFFF"),
        "36" => String::from("#800080"),
        "37" => String::from("#800000"),
        "38" => String::from("#008080"),
        "39" => String::from("#0000FF"),
        "40" => String::from("#00CCFF"),
        "41" => String::from("#CCFFFF"),
        "42" => String::from("#CCFFCC"),
        "43" => String::from("#FFFF99"),
        "44" => String::from("#99CCFF"),
        "45" => String::from("#FF99CC"),
        "46" => String::from("#CC99FF"),
        "47" => String::from("#FFCC99"),
        "48" => String::from("#3366FF"),
        "49" => String::from("#33CCCC"),
        "50" => String::from("#99CC00"),
        "51" => String::from("#FFCC00"),
        "52" => String::from("#FF9900"),
        "53" => String::from("#FF6600"),
        "54" => String::from("#666699"),
        "55" => String::from("#969696"),
        "56" => String::from("#003366"),
        "57" => String::from("#339966"),
        "58" => String::from("#003300"),
        "59" => String::from("#333300"),
        "60" => String::from("#993300"),
        "61" => String::from("#993366"),
        "62" => String::from("#333399"),
        "63" => String::from("#333333"),
        _ => String::from("#000000")
    }
}

fn get_format(code: &str) -> Option<String> {
    match code {
        "0" => Some(String::from("General")),
        "1" => Some(String::from("0")),
        "2" => Some(String::from("0.00")),
        "3" => Some(String::from("#,##0")),
        "4" => Some(String::from("#,##0.00")),
        "9" => Some(String::from("0%")),
        "10" => Some(String::from("0.00%")),
        "11" => Some(String::from("0.00E+00")),
        "12" => Some(String::from("# ?/?")),
        "13" => Some(String::from("# ??/??")),
        "14" => Some(String::from("mm-dd-yy")),
        "15" => Some(String::from("d-mmm-yy")),
        "16" => Some(String::from("d-mmm")),
        "17" => Some(String::from("mmm-yy")),
        "18" => Some(String::from("h:mm AM/PM")),
        "19" => Some(String::from("h:mm:ss AM/PM")),
        "20" => Some(String::from("h:mm")),
        "21" => Some(String::from("h:mm:ss")),
        "22" => Some(String::from("m/d/yy h:mm")),
        "37" => Some(String::from("#,##0 ;(#,##0)")),
        "38" => Some(String::from("#,##0 ;[Red](#,##0)")),
        "40" => Some(String::from("#,##0.00;[Red](#,##0.00)")),
        "45" => Some(String::from("mm:ss")),
        "46" => Some(String::from("[h]:mm:ss")),
        "47" => Some(String::from("mmss.0")),
        "48" => Some(String::from("##0.0E+0")),
        "49" => Some(String::from("@")),
        _ => None
    }
}


#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn file_read() {
        use std::io::Read;

        let now = std::time::Instant::now();
        {
            let mut file = std::fs::File::open("./example/file_example_XLSX_5000.xlsx").unwrap();
            let mut buf = vec!();
            file.read_to_end(&mut buf).unwrap();
            let mut xlsx = XLSX::new(buf);
            // // cant test with get_sheet_data coz it return jsValue
            let (name, path) = xlsx.sheets[0].clone();
            let _data = xlsx.read_sheet(path, name).unwrap();
            let _styles = xlsx.read_style();
        }
        let elapsed = now.elapsed();
        let sec = (elapsed.as_secs() as f64) + (elapsed.subsec_nanos() as f64 / 1000_000_000.0);
        println!("time to read 5000 rows: {}",  sec);
    }

    #[test]
    fn cell_to_offsets_test() {
        assert_eq!(cell_index_to_offsets(String::from("A24")), (0, 23));
        assert_eq!(cell_index_to_offsets(String::from("AB1")), (27, 0));
        assert_eq!(cell_index_to_offsets(String::from("ZZ100")), (701, 99));
    }
}
