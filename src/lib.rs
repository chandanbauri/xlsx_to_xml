use calamine::{open_workbook, Reader, Xlsx};
use pyo3::exceptions::PyRuntimeError;
use pyo3::prelude::*;
use quick_xml::events::{BytesEnd, BytesStart, Event,BytesText};
use quick_xml::Writer;
use std::io::Cursor;

#[pyfunction]
fn convert_xlsx_to_xml(file_path: &str,sheet_name: &str) -> PyResult<String> {
    let data = read_xlsx(file_path,sheet_name).map_err(|e| {
        PyRuntimeError::new_err(format!("Failed to read XLSX file: {}", e))
    })?;
    
    let xml = write_xml(data).map_err(|e| {
        PyRuntimeError::new_err(format!("Failed to write XML: {}", e))
    })?;
    
    Ok(xml)
}

fn read_xlsx(file_path: &str,sheet_name: &str) -> Result<Vec<Vec<String>>, Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook(file_path)?;
    let mut data = Vec::new();

    if let Some(Ok(sheet)) = workbook.worksheet_range(sheet_name) {
        for row in sheet.rows() {
            let row_data: Vec<String> = row.iter().map(|cell| cell.to_string()).collect();
            data.push(row_data);
        }
    } else {
        return Err("Sheet1 not found or could not be read".into());
    }

    Ok(data)
}


fn write_xml(data: Vec<Vec<String>>) -> Result<String, Box<dyn std::error::Error>> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));

    let root = BytesStart::new("Workbook");
    writer.write_event(Event::Start(root))?;

    for row in data {
        let row_start = BytesStart::new("Row");
        writer.write_event(Event::Start(row_start))?;

        for cell in row {
            let cell_start = BytesStart::new("Cell");
            writer.write_event(Event::Start(cell_start))?;
            
            // Convert `cell` to bytes and use BytesText::from
            let cell_bytes = cell;
            writer.write_event(Event::Text(BytesText::from_escaped(&cell_bytes)))?;
            
            writer.write_event(Event::End(BytesEnd::new("Cell")))?;
        }

        writer.write_event(Event::End(BytesEnd::new("Row")))?;
    }

    writer.write_event(Event::End(BytesEnd::new("Workbook")))?;

    let result = writer.into_inner().into_inner();
    let xml_string = String::from_utf8(result)?;
    Ok(xml_string)
}
/// A Python module implemented in Rust.
#[pymodule]
fn xlsx_to_xml(py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(convert_xlsx_to_xml, m)?)?;
    Ok(())
}
