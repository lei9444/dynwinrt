use dynwinrt::{MethodSignature, VTableSignature};
use pyo3::prelude::*;

#[pyclass]
#[derive(Debug, Clone)]
pub struct WinRTMethod(MethodSignature);

#[pymethods]
impl WinRTMethod {
    #[new]
    fn new() -> Self {
        Self(MethodSignature::new())
    }

    fn __repr__(&self) -> String {
        format!("{:?}", self.0)
    }
}

#[pyclass(unsendable)]
#[derive(Debug)]
pub struct WinRTInterface(VTableSignature);

#[pymethods]
impl WinRTInterface {
    #[new]
    fn new() -> Self {
        Self(VTableSignature::new())
    }

    fn add_method(&mut self, method: &WinRTMethod) -> PyResult<String> {
        self.0.add_method(method.0.clone());
        Ok(format!("Added method: {:?}", method.0))
    }
}
