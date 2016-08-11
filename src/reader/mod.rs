pub mod workbook;

use std::os::raw::c_char;
use std::ffi::CStr;
use std::str;

#[link(name = "xlsreader")]
extern {
    fn xls_getVersion() -> *const c_char;
}

pub fn version() -> String {
    let c_string = unsafe { CStr::from_ptr(xls_getVersion()) };
    str::from_utf8(c_string.to_bytes()).unwrap().to_string()
}
