use pyo3::prelude::*;
use pyo3::types::{PyInt, PyFloat, PyString, PyList, PyDict};

mod values;
mod runtime;

/// A Python module implemented in Rust.
#[pymodule]
mod dynwinrt_py {
    use pyo3::prelude::*;
    use pyo3::types::{PyInt, PyFloat, PyString, PyList, PyDict};

    /// Formats the sum of two numbers as string.
    #[pyfunction]
    fn sum_as_string(a: usize, b: usize) -> PyResult<String> {
        Ok((a + b).to_string())
    }

    #[pyfunction]
    fn sum_via_interop(a: f64, b: f64) -> PyResult<f64> {
        Ok(dynwinrt::export_add(a, &b))
    }

    #[pyfunction]
    fn process_value(value: &Bound<'_, PyAny>) -> PyResult<String> {
        if let Ok(i) = value.extract::<i64>() {
            Ok(format!("Integer: {}", i))
        } else if let Ok(f) = value.extract::<f64>() {
            Ok(format!("Float: {}", f))
        } else if let Ok(s) = value.extract::<String>() {
            Ok(format!("String: {}", s))
        } else if value.downcast::<PyList>().is_ok() {
            let list = value.downcast::<PyList>()?;
            Ok(format!("List with {} items", list.len()))
        } else if value.downcast::<PyDict>().is_ok() {
            Ok("Dictionary".to_string())
        } else {
            Ok(format!("Unknown type: {}", value.get_type().name()?))
        }
    }

    #[pyfunction]
    fn check_type(value: &Bound<'_, PyAny>) -> PyResult<String> {
        if value.is_instance_of::<PyInt>() {
            Ok("int".to_string())
        } else if value.is_instance_of::<PyFloat>() {
            Ok("float".to_string())
        } else if value.is_none() {
            Ok("None".to_string())
        } else {
            Ok("other".to_string())
        }
    }

    #[pymodule_init]
    fn init(m: &Bound<'_, PyModule>) -> PyResult<()> {
        m.add_class::<super::runtime::WinRTMethod>()?;
        m.add_class::<super::runtime::WinRTInterface>()?;
        Ok(())
    }
}
