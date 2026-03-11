use std::sync::Arc;

use dynwinrt;
use pyo3::exceptions::PyRuntimeError;
use pyo3::prelude::*;
use windows::core::{GUID, HSTRING, Interface};

/// Shared MetadataTable — created once, used everywhere.
static TABLE: std::sync::LazyLock<Arc<dynwinrt::MetadataTable>> =
    std::sync::LazyLock::new(|| dynwinrt::MetadataTable::new());

// ======================================================================
// Runtime initialization
// ======================================================================

#[pyclass]
pub struct WinAppSDKContext(#[allow(dead_code)] dynwinrt::WinAppSdkContext);

#[pyfunction]
pub fn init_winappsdk(major: u32, minor: u32) -> PyResult<WinAppSDKContext> {
    dynwinrt::initialize_winappsdk(major, minor)
        .map(WinAppSDKContext)
        .map_err(|e| PyRuntimeError::new_err(e.message()))
}

#[pyfunction]
pub fn ro_initialize(apartment_type: Option<i32>) -> PyResult<()> {
    use windows::Win32::System::WinRT::{RoInitialize, RO_INIT_MULTITHREADED, RO_INIT_SINGLETHREADED};
    let init_type = match apartment_type.unwrap_or(1) {
        0 => RO_INIT_SINGLETHREADED,
        _ => RO_INIT_MULTITHREADED,
    };
    unsafe { RoInitialize(init_type) }
        .map_err(|e| PyRuntimeError::new_err(e.message()))
}

#[pyfunction]
pub fn ro_uninitialize() {
    use windows::Win32::System::WinRT::RoUninitialize;
    unsafe { RoUninitialize() };
}

// ======================================================================
// WinGUID
// ======================================================================

#[pyclass]
#[derive(Debug, Clone, Copy)]
pub struct WinGUID(pub(crate) GUID);

#[pymethods]
impl WinGUID {
    #[staticmethod]
    fn parse(guid_str: &str) -> PyResult<Self> {
        let guid = GUID::try_from(guid_str)
            .map_err(|e| PyRuntimeError::new_err(format!("Invalid GUID: {:?}", e)))?;
        Ok(WinGUID(guid))
    }

    fn __repr__(&self) -> String {
        format!("WinGUID({:?})", self.0)
    }
}

// ======================================================================
// DynWinRTType — wraps TypeHandle
// ======================================================================

#[pyclass]
#[derive(Clone)]
pub struct DynWinRTType(dynwinrt::TypeHandle);

#[pymethods]
impl DynWinRTType {
    // -- Primitive types --

    #[staticmethod]
    fn i32_type() -> Self {
        DynWinRTType(TABLE.i32_type())
    }

    #[staticmethod]
    fn i64_type() -> Self {
        DynWinRTType(TABLE.i64_type())
    }

    #[staticmethod]
    fn hstring() -> Self {
        DynWinRTType(TABLE.hstring())
    }

    #[staticmethod]
    fn object() -> Self {
        DynWinRTType(TABLE.object())
    }

    #[staticmethod]
    fn f64_type() -> Self {
        DynWinRTType(TABLE.f64_type())
    }

    #[staticmethod]
    fn f32_type() -> Self {
        DynWinRTType(TABLE.f32_type())
    }

    #[staticmethod]
    fn u8_type() -> Self {
        DynWinRTType(TABLE.u8_type())
    }

    #[staticmethod]
    fn u16_type() -> Self {
        DynWinRTType(TABLE.u16_type())
    }

    #[staticmethod]
    fn u32_type() -> Self {
        DynWinRTType(TABLE.u32_type())
    }

    #[staticmethod]
    fn u64_type() -> Self {
        DynWinRTType(TABLE.u64_type())
    }

    #[staticmethod]
    fn i8_type() -> Self {
        DynWinRTType(TABLE.i8_type())
    }

    #[staticmethod]
    fn i16_type() -> Self {
        DynWinRTType(TABLE.i16_type())
    }

    #[staticmethod]
    fn bool_type() -> Self {
        DynWinRTType(TABLE.bool_type())
    }

    // -- Class / interface types --

    #[staticmethod]
    fn runtime_class(name: String, default_iid: &WinGUID) -> Self {
        DynWinRTType(TABLE.runtime_class(name, default_iid.0))
    }

    #[staticmethod]
    fn interface(iid: &WinGUID) -> Self {
        DynWinRTType(TABLE.interface(iid.0))
    }

    // -- Async types --

    #[staticmethod]
    fn i_async_action() -> Self {
        DynWinRTType(TABLE.async_action())
    }

    #[staticmethod]
    fn i_async_action_with_progress(progress_type: &DynWinRTType) -> Self {
        DynWinRTType(TABLE.async_action_with_progress(&progress_type.0))
    }

    #[staticmethod]
    fn i_async_operation(result_type: &DynWinRTType) -> Self {
        DynWinRTType(TABLE.async_operation(&result_type.0))
    }

    #[staticmethod]
    fn i_async_operation_with_progress(
        result_type: &DynWinRTType,
        progress_type: &DynWinRTType,
    ) -> Self {
        DynWinRTType(TABLE.async_operation_with_progress(&result_type.0, &progress_type.0))
    }

    // -- Composite types --

    #[staticmethod]
    fn struct_type(fields: Vec<DynWinRTType>) -> Self {
        let handles: Vec<dynwinrt::TypeHandle> = fields.iter().map(|f| f.0.clone()).collect();
        DynWinRTType(TABLE.define_struct(&handles))
    }

    #[staticmethod]
    fn register_struct(name: String, fields: Vec<DynWinRTType>) -> Self {
        let handles: Vec<dynwinrt::TypeHandle> = fields.iter().map(|f| f.0.clone()).collect();
        DynWinRTType(TABLE.define_named_struct(&name, &handles))
    }

    #[staticmethod]
    fn named_enum(name: String) -> Self {
        DynWinRTType(TABLE.define_named_enum(&name))
    }

    #[staticmethod]
    fn parameterized(generic_iid: &WinGUID, args: Vec<DynWinRTType>) -> Self {
        let handles: Vec<dynwinrt::TypeHandle> = args.iter().map(|a| a.0.clone()).collect();
        let generic = TABLE.generic(generic_iid.0, handles.len() as u32);
        DynWinRTType(TABLE.parameterized(&generic, &handles))
    }

    #[staticmethod]
    fn array_type(element_type: &DynWinRTType) -> Self {
        DynWinRTType(TABLE.array(&element_type.0))
    }

    // -- Interface registration & method management --

    #[staticmethod]
    fn register_interface(name: String, iid: &WinGUID) -> Self {
        DynWinRTType(TABLE.register_interface(&name, iid.0))
    }

    /// Add a method to this interface. Returns new DynWinRTType for chaining.
    fn add_method(&self, name: String, sig: &DynWinRTMethodSig) -> DynWinRTType {
        DynWinRTType(self.0.clone().add_method(&name, sig.0.clone()))
    }

    /// Get a MethodHandle by vtable index (6 = first user method).
    fn method(&self, vtable_index: usize) -> PyResult<DynWinRTMethodHandle> {
        self.0
            .method(vtable_index)
            .map(DynWinRTMethodHandle)
            .ok_or_else(|| {
                PyRuntimeError::new_err(format!("No method at vtable index {}", vtable_index))
            })
    }

    /// Get a MethodHandle by method name.
    fn method_by_name(&self, name: &str) -> PyResult<DynWinRTMethodHandle> {
        self.0
            .method_by_name(name)
            .map(DynWinRTMethodHandle)
            .ok_or_else(|| PyRuntimeError::new_err(format!("Method '{}' not found", name)))
    }

    fn __repr__(&self) -> String {
        format!("DynWinRTType({:?})", self.0.kind())
    }
}

// ======================================================================
// DynWinRTMethodSig — builder for method parameter descriptions
// ======================================================================

#[pyclass]
#[derive(Clone)]
pub struct DynWinRTMethodSig(pub(crate) dynwinrt::MethodSignature);

#[pymethods]
impl DynWinRTMethodSig {
    #[new]
    fn new() -> Self {
        DynWinRTMethodSig(dynwinrt::MethodSignature::new(&*TABLE))
    }

    /// Add an [in] parameter. Returns new sig for chaining.
    fn add_in(&self, typ: &DynWinRTType) -> DynWinRTMethodSig {
        DynWinRTMethodSig(self.0.clone().add_in(typ.0.clone()))
    }

    /// Add an [out] parameter. Returns new sig for chaining.
    fn add_out(&self, typ: &DynWinRTType) -> DynWinRTMethodSig {
        DynWinRTMethodSig(self.0.clone().add_out(typ.0.clone()))
    }
}

// ======================================================================
// DynWinRTMethodHandle — method invocation wrapper
// ======================================================================

#[pyclass]
pub struct DynWinRTMethodHandle(dynwinrt::MethodHandle);

#[pymethods]
impl DynWinRTMethodHandle {
    /// Invoke this method on a COM object.
    fn invoke(&self, obj: &DynWinRTValue, args: Vec<DynWinRTValue>) -> PyResult<DynWinRTValue> {
        let raw = match &obj.0 {
            dynwinrt::WinRTValue::Object(o) => o.as_raw(),
            _ => return Err(PyRuntimeError::new_err("invoke() requires an Object value")),
        };
        let wrt_args: Vec<dynwinrt::WinRTValue> = args.iter().map(|a| a.0.clone()).collect();
        let results = self
            .0
            .invoke(raw, &wrt_args)
            .map_err(|e| PyRuntimeError::new_err(e.message()))?;
        if results.is_empty() {
            Ok(DynWinRTValue(dynwinrt::WinRTValue::I32(0)))
        } else {
            Ok(DynWinRTValue(results.into_iter().next().unwrap()))
        }
    }
}

// ======================================================================
// DynWinRTValue — main value container
// ======================================================================

#[pyclass]
#[derive(Clone)]
pub struct DynWinRTValue(pub(crate) dynwinrt::WinRTValue);

#[pymethods]
impl DynWinRTValue {
    #[staticmethod]
    fn activation_factory(name: String) -> PyResult<DynWinRTValue> {
        dynwinrt::ro_get_activation_factory_2(&HSTRING::from(name))
            .map(DynWinRTValue)
            .map_err(|e| PyRuntimeError::new_err(e.message()))
    }

    #[staticmethod]
    fn from_i64(value: i64) -> DynWinRTValue {
        DynWinRTValue(dynwinrt::WinRTValue::I64(value))
    }

    #[staticmethod]
    fn from_i32(value: i32) -> DynWinRTValue {
        DynWinRTValue(dynwinrt::WinRTValue::I32(value))
    }

    #[staticmethod]
    fn from_u32(value: u32) -> DynWinRTValue {
        DynWinRTValue(dynwinrt::WinRTValue::U32(value))
    }

    #[staticmethod]
    fn from_f32(value: f32) -> DynWinRTValue {
        DynWinRTValue(dynwinrt::WinRTValue::F32(value))
    }

    #[staticmethod]
    fn from_f64(value: f64) -> DynWinRTValue {
        DynWinRTValue(dynwinrt::WinRTValue::F64(value))
    }

    #[staticmethod]
    fn from_bool(value: bool) -> DynWinRTValue {
        DynWinRTValue(dynwinrt::WinRTValue::Bool(value))
    }

    #[staticmethod]
    fn from_hstring(value: String) -> DynWinRTValue {
        DynWinRTValue(dynwinrt::WinRTValue::HString(HSTRING::from(value)))
    }

    /// Await an async WinRT operation (blocks the current thread).
    fn wait(&self) -> PyResult<DynWinRTValue> {
        let v = pollster::block_on(async { (&self.0).await })
            .map_err(|e| PyRuntimeError::new_err(e.message()))?;
        Ok(DynWinRTValue(v))
    }

    fn to_string(&self) -> String {
        match &self.0 {
            dynwinrt::WinRTValue::HString(s) => s.to_string(),
            dynwinrt::WinRTValue::I32(i) => i.to_string(),
            dynwinrt::WinRTValue::I64(i) => i.to_string(),
            dynwinrt::WinRTValue::U32(i) => i.to_string(),
            dynwinrt::WinRTValue::U64(i) => i.to_string(),
            dynwinrt::WinRTValue::F32(f) => f.to_string(),
            dynwinrt::WinRTValue::F64(f) => f.to_string(),
            dynwinrt::WinRTValue::Bool(b) => b.to_string(),
            dynwinrt::WinRTValue::Object(o) => format!("Object({:?})", o),
            _ => "Unsupported type".to_string(),
        }
    }

    fn __repr__(&self) -> String {
        format!("DynWinRTValue({})", self.to_string())
    }

    fn __str__(&self) -> String {
        self.to_string()
    }

    fn to_int(&self) -> PyResult<i64> {
        match &self.0 {
            dynwinrt::WinRTValue::I32(i) => Ok(*i as i64),
            dynwinrt::WinRTValue::I64(i) => Ok(*i),
            dynwinrt::WinRTValue::U32(i) => Ok(*i as i64),
            dynwinrt::WinRTValue::U64(i) => Ok(*i as i64),
            dynwinrt::WinRTValue::Bool(b) => Ok(*b as i64),
            _ => Err(PyRuntimeError::new_err("Cannot convert to int")),
        }
    }

    fn to_float(&self) -> PyResult<f64> {
        match &self.0 {
            dynwinrt::WinRTValue::F32(f) => Ok(*f as f64),
            dynwinrt::WinRTValue::F64(f) => Ok(*f),
            dynwinrt::WinRTValue::I32(i) => Ok(*i as f64),
            _ => Err(PyRuntimeError::new_err("Cannot convert to float")),
        }
    }

    fn as_raw(&self) -> PyResult<i64> {
        match &self.0 {
            dynwinrt::WinRTValue::Object(o) => Ok(o.as_raw() as i64),
            _ => Err(PyRuntimeError::new_err(
                "Cannot get raw pointer from non-object",
            )),
        }
    }

    /// COM QueryInterface — cast to a different interface.
    fn cast(&self, iid: &WinGUID) -> PyResult<DynWinRTValue> {
        self.0
            .cast(&iid.0)
            .map(DynWinRTValue)
            .map_err(|e| PyRuntimeError::new_err(e.message()))
    }

    // -- Convenience call methods (match JS API) --

    /// Call a method with no args and one out param.
    fn call_0(&self, method_index: usize, return_type: &DynWinRTType) -> PyResult<DynWinRTValue> {
        let method = dynwinrt::MethodSignature::new(&*TABLE)
            .add_out(return_type.0.clone())
            .build(method_index);
        let obj_raw = self
            .0
            .as_object()
            .ok_or_else(|| PyRuntimeError::new_err("call_0 requires an Object value"))?
            .as_raw();
        let result = method
            .call_dynamic(obj_raw, &[])
            .map_err(|e| PyRuntimeError::new_err(e.message()))?;
        Ok(DynWinRTValue(result.into_iter().next().unwrap()))
    }

    /// Call a method with one arg and one out param.
    fn call_1(
        &self,
        method_index: usize,
        return_type: &DynWinRTType,
        v1: &DynWinRTValue,
    ) -> PyResult<DynWinRTValue> {
        let in_type = TABLE.handle_from_kind(v1.0.get_type_kind());
        let method = dynwinrt::MethodSignature::new(&*TABLE)
            .add_in(in_type)
            .add_out(return_type.0.clone())
            .build(method_index);
        let obj_raw = self
            .0
            .as_object()
            .ok_or_else(|| PyRuntimeError::new_err("call_1 requires an Object value"))?
            .as_raw();
        let result = method
            .call_dynamic(obj_raw, &[v1.0.clone()])
            .map_err(|e| PyRuntimeError::new_err(e.message()))?;
        Ok(DynWinRTValue(result.into_iter().next().unwrap()))
    }

    /// General-purpose method call with explicit types and args.
    fn call(
        &self,
        method_index: usize,
        return_type: &DynWinRTType,
        in_types: Vec<DynWinRTType>,
        args: Vec<DynWinRTValue>,
    ) -> PyResult<DynWinRTValue> {
        let mut method = dynwinrt::MethodSignature::new(&*TABLE);
        for t in &in_types {
            method = method.add_in(t.0.clone());
        }
        method = method.add_out(return_type.0.clone());

        let obj = match &self.0 {
            dynwinrt::WinRTValue::Object(o) => o.as_raw(),
            _ => return Err(PyRuntimeError::new_err("call() requires an Object value")),
        };

        let mut iface = dynwinrt::InterfaceSignature::define_from_iinspectable(
            "",
            Default::default(),
            &*TABLE,
        );
        let target_index = method_index;
        for _ in 6..target_index {
            iface.add_method(dynwinrt::MethodSignature::new(&*TABLE));
        }
        iface.add_method(method);

        let winrt_args: Vec<dynwinrt::WinRTValue> = args.iter().map(|a| a.0.clone()).collect();
        let result = iface.methods[target_index]
            .call_dynamic(obj, &winrt_args)
            .map_err(|e| PyRuntimeError::new_err(e.message()))?;

        if result.is_empty() {
            Ok(DynWinRTValue(dynwinrt::WinRTValue::I32(0)))
        } else {
            Ok(DynWinRTValue(result.into_iter().next().unwrap()))
        }
    }

    // -- Array / Struct extraction --

    fn is_array(&self) -> bool {
        self.0.as_array().is_some()
    }

    fn as_array(&self) -> PyResult<DynWinRTArray> {
        match &self.0 {
            dynwinrt::WinRTValue::Array(data) => Ok(DynWinRTArray(data.clone())),
            _ => Err(PyRuntimeError::new_err("Value is not an Array")),
        }
    }

    fn is_struct(&self) -> bool {
        self.0.as_struct().is_some()
    }

    fn as_struct(&self) -> PyResult<DynWinRTStruct> {
        match &self.0 {
            dynwinrt::WinRTValue::Struct(data) => Ok(DynWinRTStruct(data.clone())),
            _ => Err(PyRuntimeError::new_err("Value is not a Struct")),
        }
    }
}

// ======================================================================
// DynWinRTArray — array container with blittable fast paths
// ======================================================================

#[pyclass(unsendable)]
#[derive(Clone)]
pub struct DynWinRTArray(dynwinrt::ArrayData);

#[pymethods]
impl DynWinRTArray {
    fn __len__(&self) -> usize {
        self.0.len()
    }

    /// Per-element access.
    fn get(&self, index: usize) -> DynWinRTValue {
        DynWinRTValue(self.0.get(index))
    }

    /// Convert all elements to a list of DynWinRTValue.
    fn to_values(&self) -> Vec<DynWinRTValue> {
        (0..self.0.len())
            .map(|i| DynWinRTValue(self.0.get(i)))
            .collect()
    }

    // -- Blittable fast paths --

    fn to_i32_list(&self) -> Vec<i32> {
        unsafe { self.0.as_typed_slice::<i32>().to_vec() }
    }

    fn to_u32_list(&self) -> Vec<u32> {
        unsafe { self.0.as_typed_slice::<u32>().to_vec() }
    }

    fn to_f32_list(&self) -> Vec<f32> {
        unsafe { self.0.as_typed_slice::<f32>().to_vec() }
    }

    fn to_f64_list(&self) -> Vec<f64> {
        unsafe { self.0.as_typed_slice::<f64>().to_vec() }
    }

    fn to_u8_list(&self) -> Vec<u8> {
        unsafe { self.0.as_typed_slice::<u8>().to_vec() }
    }

    fn to_i64_list(&self) -> Vec<i64> {
        unsafe { self.0.as_typed_slice::<i64>().to_vec() }
    }

    // -- Construction from Python lists --

    #[staticmethod]
    fn from_i32_values(values: Vec<i32>) -> DynWinRTArray {
        let wvals: Vec<dynwinrt::WinRTValue> =
            values.into_iter().map(dynwinrt::WinRTValue::I32).collect();
        DynWinRTArray(dynwinrt::ArrayData::from_values(TABLE.i32_type(), &wvals))
    }

    #[staticmethod]
    fn from_f64_values(values: Vec<f64>) -> DynWinRTArray {
        let wvals: Vec<dynwinrt::WinRTValue> =
            values.into_iter().map(dynwinrt::WinRTValue::F64).collect();
        DynWinRTArray(dynwinrt::ArrayData::from_values(TABLE.f64_type(), &wvals))
    }

    #[staticmethod]
    fn from_u8_values(values: Vec<u8>) -> DynWinRTArray {
        let wvals: Vec<dynwinrt::WinRTValue> =
            values.into_iter().map(dynwinrt::WinRTValue::U8).collect();
        DynWinRTArray(dynwinrt::ArrayData::from_values(TABLE.u8_type(), &wvals))
    }

    /// Wrap as DynWinRTValue::Array for passing to call().
    fn to_value(&self) -> DynWinRTValue {
        DynWinRTValue(dynwinrt::WinRTValue::Array(self.0.clone()))
    }

    fn __repr__(&self) -> String {
        format!("DynWinRTArray(len={})", self.0.len())
    }
}

// ======================================================================
// DynWinRTStruct — typed field access by index
// ======================================================================

#[pyclass(unsendable)]
#[derive(Clone)]
pub struct DynWinRTStruct(dynwinrt::ValueTypeData);

#[pymethods]
impl DynWinRTStruct {
    /// Create a zero-initialized struct of the given type.
    #[staticmethod]
    fn create(typ: &DynWinRTType) -> DynWinRTStruct {
        DynWinRTStruct(typ.0.default_value())
    }

    fn get_i32(&self, index: usize) -> i32 {
        self.0.get_field::<i32>(index)
    }
    fn set_i32(&mut self, index: usize, value: i32) {
        self.0.set_field(index, value);
    }

    fn get_u32(&self, index: usize) -> u32 {
        self.0.get_field::<u32>(index)
    }
    fn set_u32(&mut self, index: usize, value: u32) {
        self.0.set_field(index, value);
    }

    fn get_f32(&self, index: usize) -> f32 {
        self.0.get_field::<f32>(index)
    }
    fn set_f32(&mut self, index: usize, value: f32) {
        self.0.set_field(index, value);
    }

    fn get_f64(&self, index: usize) -> f64 {
        self.0.get_field::<f64>(index)
    }
    fn set_f64(&mut self, index: usize, value: f64) {
        self.0.set_field(index, value);
    }

    fn get_i64(&self, index: usize) -> i64 {
        self.0.get_field::<i64>(index)
    }
    fn set_i64(&mut self, index: usize, value: i64) {
        self.0.set_field(index, value);
    }

    fn get_u8(&self, index: usize) -> u8 {
        self.0.get_field::<u8>(index)
    }
    fn set_u8(&mut self, index: usize, value: u8) {
        self.0.set_field(index, value);
    }

    /// Wrap as DynWinRTValue::Struct for passing to call().
    fn to_value(&self) -> DynWinRTValue {
        DynWinRTValue(dynwinrt::WinRTValue::Struct(self.0.clone()))
    }

    fn __repr__(&self) -> String {
        "DynWinRTStruct(...)".to_string()
    }
}
