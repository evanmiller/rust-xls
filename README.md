# rust-xls
Read Excel files from Rust. Requires **[libxls](http://libxls.sourceforge.net)**.

Example program:

```
extern crate xls;

use std::env;

fn main() {
    let version = xls::reader::version();
    let arg_vec = env::args().collect::<Vec<String>>();
    if arg_vec.len() == 2 {
        println!("libxls version: {}", version);
        let maybe_handle = xls::reader::workbook::new(arg_vec[1].clone());
        match maybe_handle {
            Some(workbook) => {
                println!("Success!");
                for sheet in workbook.sheets() {
                    println!("+ Sheet");
                    for row in sheet.rows() {
                        println!("++ Row {}", row.index);
                        for cell in row.cells() {
                            match cell.value() {
                                Some(value) =>  {
                                    println!("+++ Cell {}: {}", cell.col_number,
                                             value);
                                },
                                None => { },
                            }
                        }
                    }
                }
            },
            None => println!("Failure!"),
        }
    } else {
        println!("Usage: {} <filename>", arg_vec[0]);
        std::process::exit(1);
    };
}
```
