use libc::c_char;
use libc::types::common::c95::c_void;
use libc::types::os::arch::c95::c_int;
use std::mem;
use std::ffi::CString;
use std::ffi::CStr;
use std::str;
use std::fmt::Display;
use std::fmt::Formatter;
use std::fmt::Result;

pub struct Worksheet {
    pointer : *const c_void,
}
pub struct Worksheets<'a> {
    workbook : &'a Workbook,
    current_index : usize,
}

pub struct Workbook {
    pointer : *const c_void,
}

pub struct Row<'a> {
    pub index: i32,
    worksheet: &'a Worksheet,
    pointer: *const c_void,
} 
pub struct Rows<'a> {
    worksheet : &'a Worksheet,
    current_row : i32,
}

#[repr(C)]
struct NativeCell {
    id: u16,
    row: u16,
    col: u16,
    xf: u16,
    str: *const i8,
    d: f64,
    l: i32,
    width: u16,
    colspan: u16,
    rowspan: u16,
    is_hidden: i8,
}

enum CellType {
    ExcelRecordRK = 0x027E,
    ExcelRecordMulRK= 0x00BD,
    ExcelRecordNumber = 0x0203,
    ExcelRecordBoolErr = 0x0205,
    ExcelRecordFormula = 0x0006,
}

pub enum CellValue {
    DoubleValue(f64),
    StringValue{ value: String },
}

pub struct Cell {
    pub row_number: i32,
    pub col_number: i32,
    pointer: *const NativeCell,
}
pub struct Cells<'a> {
    row: &'a Row<'a>,
    current_col: i32,
}

extern {
    fn xls_open(filename: *const c_char, encoding: *const c_char) -> *const c_void;
    fn xls_close_WB(wb: *const c_void);
    fn xls_getWorkSheet(wb: *const c_void, index: c_int) -> *const c_void;
    fn xls_parseWorkSheet(ws: *const c_void);
    fn xls_row(ws: *const c_void, row: u16) -> *const c_void;
    fn xls_cell(ws: *const c_void, row: u16, col: u16) -> *const NativeCell;
}

fn is_null(handle: *const c_void) -> bool {
    let ptr_as_int : usize = unsafe { mem::transmute(handle) };
    return ptr_as_int == 0;
}

pub fn new(filename: String) -> Option<Workbook> {
    let file_string = CString::new(filename.clone()).unwrap();
    let enc_string = CString::new("UTF-8").unwrap();
    let handle = unsafe { xls_open(file_string.as_ptr(), enc_string.as_ptr()) };
    if !is_null(handle as *const c_void) {
        Some(Workbook{ pointer: handle })
    } else {
        None
    }
}

impl Drop for Workbook {
    fn drop(&mut self) {
        unsafe { xls_close_WB(self.pointer) };
        self.pointer = unsafe { mem::transmute(0 as usize) };
    }
}

impl Workbook {
    pub fn sheets(&self) -> Worksheets {
        Worksheets { workbook: self, current_index: 0 }
    }
}

impl<'a> Iterator for Worksheets<'a> {
    type Item = Worksheet;

    fn next(&mut self) -> Option<Worksheet> {
        let ws = unsafe { xls_getWorkSheet(self.workbook.pointer, self.current_index as c_int) };
        if is_null(ws as *const c_void) {
            None
        } else {
            self.current_index += 1;
            unsafe { xls_parseWorkSheet(ws) };
            Some(Worksheet{ pointer: ws })
        }
    }
}

impl Worksheet {
    pub fn rows(&self) -> Rows {
        Rows { worksheet: self, current_row: 0 }
    }
}


impl<'a> Iterator for Rows<'a> {
    type Item = Row<'a>;

    fn next(&mut self) -> Option<Row<'a>> {
        let row = unsafe { xls_row(self.worksheet.pointer, self.current_row as u16) };
        if is_null(row) {
            None
        } else {
            let row_struct = Row{ worksheet: self.worksheet, pointer: row, index: self.current_row };
            self.current_row += 1;
            Some(row_struct)
        }
    }
}

impl<'a> Row<'a> {
    pub fn cells(&self) -> Cells {
        Cells { row: self, current_col: 0 }
    }
}

impl<'a> Iterator for Cells<'a> {
    type Item = Cell;

    fn next(&mut self) -> Option<Cell> {
        let cell = unsafe { xls_cell(self.row.worksheet.pointer, 
                                     self.row.index as u16, 
                                     self.current_col as u16) };
        if is_null(cell as *const c_void) {
            None
        } else {
            let cell_struct = Cell{ 
                pointer : cell, 
                row_number: self.row.index, 
                col_number: self.current_col };
            self.current_col += 1;
            Some(cell_struct)
        }
    }
}

impl Cell {
    pub fn value(&self) -> Option<CellValue> {
        let id = unsafe { (*self.pointer).id };
        let str_val = unsafe { (*self.pointer).str };
        let d_val = unsafe { (*self.pointer).d };
        match id {
            id if (id == CellType::ExcelRecordNumber as u16 ||
                   id == CellType::ExcelRecordRK as u16 ||
                   id == CellType::ExcelRecordMulRK as u16 ||
                   id == CellType::ExcelRecordBoolErr as u16 ||
                   id == CellType::ExcelRecordFormula as u16) => Some(CellValue::DoubleValue(d_val)),
            _ => {
                if is_null(str_val as *const c_void) {
                    None
                } else {
                    let c_string = unsafe { CStr::from_ptr(str_val) };
                    let string_value = str::from_utf8(c_string.to_bytes()).unwrap().to_string();
                    Some(CellValue::StringValue{ value: string_value })
                }
            },
        }
    }
}

impl Display for CellValue {
    fn fmt(&self, formatter: &mut Formatter) -> Result {
        match self {
            &CellValue::DoubleValue(val) => val.fmt(formatter),
            &CellValue::StringValue{ value: ref val } => val.fmt(formatter),
        }
    }
}
