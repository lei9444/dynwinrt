use pyo3::prelude::*;

mod runtime;
mod values;

#[pymodule]
mod dynwinrt_py {
    use pyo3::prelude::*;

    #[pymodule_init]
    fn init(m: &Bound<'_, PyModule>) -> PyResult<()> {
        // Classes
        m.add_class::<super::runtime::WinAppSDKContext>()?;
        m.add_class::<super::runtime::WinGUID>()?;
        m.add_class::<super::runtime::DynWinRTType>()?;
        m.add_class::<super::runtime::DynWinRTMethodSig>()?;
        m.add_class::<super::runtime::DynWinRTMethodHandle>()?;
        m.add_class::<super::runtime::DynWinRTValue>()?;
        m.add_class::<super::runtime::DynWinRTArray>()?;
        m.add_class::<super::runtime::DynWinRTStruct>()?;

        // Functions
        m.add_function(wrap_pyfunction!(super::runtime::init_winappsdk, m)?)?;
        m.add_function(wrap_pyfunction!(super::runtime::ro_initialize, m)?)?;
        m.add_function(wrap_pyfunction!(super::runtime::ro_uninitialize, m)?)?;

        Ok(())
    }
}
