#![deny(clippy::all)]

use std::sync::Arc;

use dynwinrt;
use napi::{
  bindgen_prelude::Result,
  Error, Status,
};
use napi_derive::napi;
use windows::core::{h, IUnknown, Interface, HSTRING};
use windows_future::IAsyncOperationWithProgress;

/// Shared MetadataTable — created once, used everywhere.
static TABLE: std::sync::LazyLock<Arc<dynwinrt::MetadataTable>> =
  std::sync::LazyLock::new(|| dynwinrt::MetadataTable::new());

// ======================================================================
// Runtime initialization
// ======================================================================

#[napi]
struct WinAppSDKContext(dynwinrt::WinAppSdkContext);

#[napi]
pub fn init_winappsdk(major: u32, minor: u32) {
  WinAppSDKContext(dynwinrt::initialize_winappsdk(major, minor).unwrap());
}

#[napi]
pub fn ro_initialize(apartment_type: Option<i32>) {
  use windows::Win32::System::WinRT::{RoInitialize, RO_INIT_MULTITHREADED, RO_INIT_SINGLETHREADED};
  let init_type = match apartment_type.unwrap_or(1) {
    0 => RO_INIT_SINGLETHREADED,
    _ => RO_INIT_MULTITHREADED,
  };
  unsafe { RoInitialize(init_type).unwrap() };
}

#[napi]
pub fn ro_uninitialize() {
  use windows::Win32::System::WinRT::RoUninitialize;
  unsafe { RoUninitialize() };
}

#[napi]
pub async fn winrt_value_to_promise(x: &DynWinRTValue) -> DynWinRTValue {
  let v = (&x.0).await.unwrap();
  DynWinRTValue(v)
}

// ======================================================================
// Core types — DynWinRTType, DynWinRTMethodSig, DynWinRTMethodHandle, WinGUID
// ======================================================================

#[napi]
pub struct DynWinRTType(dynwinrt::TypeHandle);

#[napi]
impl DynWinRTType {
  #[napi]
  pub fn i32() -> Self {
    DynWinRTType(TABLE.i32_type())
  }

  #[napi]
  pub fn i64() -> Self {
    DynWinRTType(TABLE.i64_type())
  }

  #[napi]
  pub fn hstring() -> Self {
    DynWinRTType(TABLE.hstring())
  }

  #[napi]
  pub fn object() -> Self {
    DynWinRTType(TABLE.object())
  }

  #[napi]
  pub fn f64() -> Self {
    DynWinRTType(TABLE.f64_type())
  }

  #[napi]
  pub fn f32() -> Self {
    DynWinRTType(TABLE.f32_type())
  }

  #[napi]
  pub fn u8() -> Self {
    DynWinRTType(TABLE.u8_type())
  }

  #[napi]
  pub fn u32() -> Self {
    DynWinRTType(TABLE.u32_type())
  }

  #[napi]
  pub fn u64() -> Self {
    DynWinRTType(TABLE.u64_type())
  }

  #[napi]
  pub fn i8_type() -> Self {
    DynWinRTType(TABLE.i8_type())
  }

  #[napi]
  pub fn i16() -> Self {
    DynWinRTType(TABLE.i16_type())
  }

  #[napi]
  pub fn u16() -> Self {
    DynWinRTType(TABLE.u16_type())
  }

  #[napi]
  pub fn bool_type() -> Self {
    DynWinRTType(TABLE.bool_type())
  }

  #[napi]
  pub fn runtime_class(name: String, default_iid: &WinGUID) -> Self {
    DynWinRTType(TABLE.runtime_class(name, default_iid.0))
  }

  #[napi]
  pub fn interface(iid: &WinGUID) -> Self {
    DynWinRTType(TABLE.interface(iid.0))
  }

  #[napi]
  pub fn i_async_action() -> Self {
    DynWinRTType(TABLE.async_action())
  }

  #[napi]
  pub fn i_async_action_with_progress(progress_type: &DynWinRTType) -> Self {
    DynWinRTType(TABLE.async_action_with_progress(&progress_type.0))
  }

  #[napi]
  pub fn i_async_operation(result_type: &DynWinRTType) -> Self {
    DynWinRTType(TABLE.async_operation(&result_type.0))
  }

  #[napi]
  pub fn i_async_operation_with_progress(result_type: &DynWinRTType, progress_type: &DynWinRTType) -> Self {
    DynWinRTType(TABLE.async_operation_with_progress(&result_type.0, &progress_type.0))
  }

  /// Declare a struct type from field types (anonymous).
  #[napi]
  pub fn struct_type(fields: Vec<&DynWinRTType>) -> Self {
    let handles: Vec<dynwinrt::TypeHandle> = fields.iter().map(|f| f.0.clone()).collect();
    DynWinRTType(TABLE.define_struct(&handles))
  }

  /// Declare a named struct type with WinRT full name (for correct IID signature).
  #[napi]
  pub fn register_struct(name: String, fields: Vec<&DynWinRTType>) -> Self {
    let handles: Vec<dynwinrt::TypeHandle> = fields.iter().map(|f| f.0.clone()).collect();
    DynWinRTType(TABLE.define_named_struct(&name, &handles))
  }

  /// Declare a named enum type (ABI = i32, carries name for signature).
  #[napi]
  pub fn named_enum(name: String) -> Self {
    DynWinRTType(TABLE.define_named_enum(&name))
  }

  /// Declare a parameterized type (generic instantiation, e.g. IReference<UInt64>).
  #[napi]
  pub fn parameterized(generic_iid: &WinGUID, args: Vec<&DynWinRTType>) -> Self {
    let handles: Vec<dynwinrt::TypeHandle> = args.iter().map(|a| a.0.clone()).collect();
    let generic = TABLE.generic(generic_iid.0, handles.len() as u32);
    DynWinRTType(TABLE.parameterized(&generic, &handles))
  }

  /// Declare an array-of-element type for method signatures.
  #[napi]
  pub fn array_type(element_type: &DynWinRTType) -> Self {
    DynWinRTType(TABLE.array(&element_type.0))
  }

  /// Register an interface in the MetadataTable.
  /// Returns self (Interface TypeHandle) for chaining `.addMethod()`.
  #[napi]
  pub fn register_interface(name: String, iid: &WinGUID) -> Self {
    DynWinRTType(TABLE.register_interface(&name, iid.0))
  }

  /// Add a method to this interface using a MethodSignature.
  /// Methods are numbered starting at vtable index 6.
  #[napi]
  pub fn add_method(&self, name: String, sig: &DynWinRTMethodSig) -> DynWinRTType {
    DynWinRTType(self.0.clone().add_method(&name, sig.0.clone()))
  }

  /// Get a MethodHandle by vtable index (6 = first user method).
  #[napi]
  pub fn method(&self, vtable_index: i32) -> napi::Result<DynWinRTMethodHandle> {
    self.0.method(vtable_index as usize)
      .map(DynWinRTMethodHandle)
      .ok_or_else(|| napi::Error::from_reason(
        format!("No method at vtable index {}", vtable_index)
      ))
  }

  /// Get a MethodHandle by method name.
  #[napi]
  pub fn method_by_name(&self, name: String) -> napi::Result<DynWinRTMethodHandle> {
    self.0.method_by_name(&name)
      .map(DynWinRTMethodHandle)
      .ok_or_else(|| napi::Error::from_reason(
        format!("Method '{}' not found", name)
      ))
  }
}

#[napi]
#[derive(Debug, Clone, Copy)]
pub struct WinGUID(windows::core::GUID);

#[napi]
impl WinGUID {
  #[napi]
  pub fn parse(guid_str: String) -> Self {
    let guid = windows::core::GUID::try_from(guid_str.as_str()).unwrap();
    WinGUID(guid)
  }
}

// ======================================================================
// MethodSignature binding — builder for method parameter descriptions
// ======================================================================

#[napi]
pub struct DynWinRTMethodSig(dynwinrt::MethodSignature);
unsafe impl Send for DynWinRTMethodSig {}
unsafe impl Sync for DynWinRTMethodSig {}

#[napi]
impl DynWinRTMethodSig {
  #[napi(constructor)]
  pub fn new() -> Self {
    DynWinRTMethodSig(dynwinrt::MethodSignature::new(&*TABLE))
  }

  /// Add an [in] parameter.
  #[napi]
  pub fn add_in(&self, typ: &DynWinRTType) -> DynWinRTMethodSig {
    DynWinRTMethodSig(self.0.clone().add_in(typ.0.clone()))
  }

  /// Add an [out] parameter.
  #[napi]
  pub fn add_out(&self, typ: &DynWinRTType) -> DynWinRTMethodSig {
    DynWinRTMethodSig(self.0.clone().add_out(typ.0.clone()))
  }
}

// ======================================================================
// MethodHandle binding
// ======================================================================

#[napi]
pub struct DynWinRTMethodHandle(dynwinrt::MethodHandle);
unsafe impl Send for DynWinRTMethodHandle {}
unsafe impl Sync for DynWinRTMethodHandle {}

#[napi]
impl DynWinRTMethodHandle {
  /// Invoke this method on a COM object.
  #[napi]
  pub fn invoke(
    &self,
    obj: &DynWinRTValue,
    args: Vec<&DynWinRTValue>,
  ) -> napi::Result<DynWinRTValue> {
    let raw = match &obj.0 {
      dynwinrt::WinRTValue::Object(o) => o.as_raw(),
      _ => return Err(napi::Error::from_reason("invoke() requires an Object value")),
    };
    let wrt_args: Vec<dynwinrt::WinRTValue> = args.iter().map(|a| a.0.clone()).collect();
    let results = self.0.invoke(raw, &wrt_args)
      .map_err(|e| napi::Error::from_reason(e.message()))?;
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

#[napi]
pub struct DynWinRTValue(dynwinrt::WinRTValue);
unsafe impl Send for DynWinRTValue {}
unsafe impl Sync for DynWinRTValue {}

#[napi]
impl DynWinRTValue {
  #[napi]
  pub fn activation_factory(name: String) -> DynWinRTValue {
    DynWinRTValue(dynwinrt::ro_get_activation_factory_2(&HSTRING::from(name)).unwrap())
  }

  #[napi]
  pub fn i64(value: i64) -> DynWinRTValue {
    DynWinRTValue(dynwinrt::WinRTValue::I64(value))
  }
  #[napi]
  pub fn i32(value: i32) -> DynWinRTValue {
    DynWinRTValue(dynwinrt::WinRTValue::I32(value))
  }
  #[napi]
  pub fn hstring(value: String) -> DynWinRTValue {
    DynWinRTValue(dynwinrt::WinRTValue::HString(HSTRING::from(value)))
  }

  #[napi]
  pub async fn to_promise(&self) -> DynWinRTValue {
    let v = (&self.0).await.unwrap();
    DynWinRTValue(v)
  }

  #[napi]
  pub fn to_string(&self) -> String {
    match &self.0 {
      dynwinrt::WinRTValue::HString(s) => s.to_string(),
      dynwinrt::WinRTValue::I32(i) => i.to_string(),
      dynwinrt::WinRTValue::I64(i) => i.to_string(),
      dynwinrt::WinRTValue::Object(o) => format!("Object: {:?}", o),
      _ => "Unsupported type".to_string(),
    }
  }

  #[napi]
  pub fn call_0(&self, method_index: i32, return_type: &DynWinRTType) -> DynWinRTValue {
    let method = dynwinrt::MethodSignature::new(&*TABLE)
      .add_out(return_type.0.clone())
      .build(method_index as usize);
    let obj_raw = self.0.as_object().unwrap().as_raw();
    let result = method.call_dynamic(obj_raw, &[]).unwrap();
    DynWinRTValue(result.into_iter().next().unwrap())
  }

  #[napi]
  pub fn call_single_out_0(&self, method_index: i32, return_type: &DynWinRTType) -> DynWinRTValue {
    let method = dynwinrt::MethodSignature::new(&*TABLE)
      .add_out(return_type.0.clone())
      .build(method_index as usize);
    let obj_raw = self.0.as_object().unwrap().as_raw();
    let result = method.call_dynamic(obj_raw, &[]).unwrap();
    DynWinRTValue(result.into_iter().next().unwrap())
  }

  #[napi]
  pub fn call_single_out_1(
    &self,
    method_index: i32,
    return_type: &DynWinRTType,
    v1: &DynWinRTValue,
  ) -> DynWinRTValue {
    let in_type = TABLE.handle_from_kind(v1.0.get_type_kind());
    let method = dynwinrt::MethodSignature::new(&*TABLE)
      .add_in(in_type)
      .add_out(return_type.0.clone())
      .build(method_index as usize);
    let obj_raw = self.0.as_object().unwrap().as_raw();
    let result = method.call_dynamic(obj_raw, &[v1.0.clone()]).unwrap();
    DynWinRTValue(result.into_iter().next().unwrap())
  }

  /// General-purpose method call accepting any number of arguments.
  #[napi]
  pub fn call(
    &self,
    method_index: i32,
    return_type: &DynWinRTType,
    in_types: Vec<&DynWinRTType>,
    args: Vec<&DynWinRTValue>,
  ) -> napi::Result<DynWinRTValue> {
    let mut method = dynwinrt::MethodSignature::new(&*TABLE);
    for t in &in_types {
      method = method.add_in(t.0.clone());
    }
    method = method.add_out(return_type.0.clone());

    let obj = match &self.0 {
      dynwinrt::WinRTValue::Object(o) => o.as_raw(),
      _ => return Err(napi::Error::from_reason("call() requires an Object value")),
    };

    let mut iface = dynwinrt::InterfaceSignature::define_from_iinspectable(
      "",
      Default::default(),
      &*TABLE,
    );
    let target_index = method_index as usize;
    for _ in 6..target_index {
      iface.add_method(dynwinrt::MethodSignature::new(&*TABLE));
    }
    iface.add_method(method);

    let winrt_args: Vec<dynwinrt::WinRTValue> = args.iter().map(|a| a.0.clone()).collect();
    let result = iface.methods[target_index]
      .call_dynamic(obj, &winrt_args)
      .map_err(|e| napi::Error::from_reason(e.message()))?;

    if result.is_empty() {
      Ok(DynWinRTValue(dynwinrt::WinRTValue::I32(0)))
    } else {
      Ok(DynWinRTValue(result.into_iter().next().unwrap()))
    }
  }

  #[napi]
  pub fn cast(&self, iid: &WinGUID) -> DynWinRTValue {
    let result = self.0.cast(&iid.0).unwrap();
    DynWinRTValue(result)
  }

  #[napi]
  pub fn to_number(&self) -> i32 {
    match &self.0 {
      dynwinrt::WinRTValue::I32(i) => *i,
      _ => panic!("Cannot convert to number"),
    }
  }

  #[napi]
  pub fn as_raw(&self) -> i64 {
    match &self.0 {
      dynwinrt::WinRTValue::Object(o) => o.as_raw() as i64,
      _ => panic!("Cannot get raw pointer from non-object"),
    }
  }

  // -- Array / Struct extraction --

  #[napi]
  pub fn is_array(&self) -> bool {
    self.0.as_array().is_some()
  }

  #[napi]
  pub fn as_array(&self) -> napi::Result<DynWinRTArray> {
    match &self.0 {
      dynwinrt::WinRTValue::Array(data) => Ok(DynWinRTArray(data.clone())),
      _ => Err(napi::Error::from_reason("Value is not an Array")),
    }
  }

  #[napi]
  pub fn is_struct(&self) -> bool {
    self.0.as_struct().is_some()
  }

  #[napi]
  pub fn as_struct(&self) -> napi::Result<DynWinRTStruct> {
    match &self.0 {
      dynwinrt::WinRTValue::Struct(data) => Ok(DynWinRTStruct(data.clone())),
      _ => Err(napi::Error::from_reason("Value is not a Struct")),
    }
  }
}

// ======================================================================
// Array binding — blittable fast path via typed Vec, generic fallback
// ======================================================================

#[napi]
pub struct DynWinRTArray(dynwinrt::ArrayData);
unsafe impl Send for DynWinRTArray {}
unsafe impl Sync for DynWinRTArray {}

#[napi]
impl DynWinRTArray {
  #[napi]
  pub fn len(&self) -> u32 {
    self.0.len() as u32
  }

  /// Per-element access (works for all element types).
  #[napi]
  pub fn get(&self, index: u32) -> DynWinRTValue {
    DynWinRTValue(self.0.get(index as usize))
  }

  /// Convert all elements to DynWinRTValue array.
  #[napi]
  pub fn to_values(&self) -> Vec<DynWinRTValue> {
    (0..self.0.len()).map(|i| DynWinRTValue(self.0.get(i))).collect()
  }

  // -- Blittable fast paths: zero-copy read into typed Vec --

  #[napi]
  pub fn to_i32_vec(&self) -> Vec<i32> {
    unsafe { self.0.as_typed_slice::<i32>().to_vec() }
  }

  #[napi]
  pub fn to_u32_vec(&self) -> Vec<u32> {
    unsafe { self.0.as_typed_slice::<u32>().to_vec() }
  }

  #[napi]
  pub fn to_f32_vec(&self) -> Vec<f32> {
    unsafe { self.0.as_typed_slice::<f32>().to_vec() }
  }

  #[napi]
  pub fn to_f64_vec(&self) -> Vec<f64> {
    unsafe { self.0.as_typed_slice::<f64>().to_vec() }
  }

  #[napi]
  pub fn to_u8_vec(&self) -> Vec<u8> {
    unsafe { self.0.as_typed_slice::<u8>().to_vec() }
  }

  #[napi]
  pub fn to_i64_vec(&self) -> Vec<i64> {
    unsafe { self.0.as_typed_slice::<i64>().to_vec() }
  }

  // -- Construction from JS typed arrays (blittable only) --

  #[napi]
  pub fn from_i32_values(values: Vec<i32>) -> DynWinRTArray {
    let wvals: Vec<dynwinrt::WinRTValue> = values.into_iter().map(dynwinrt::WinRTValue::I32).collect();
    DynWinRTArray(dynwinrt::ArrayData::from_values(TABLE.i32_type(), &wvals))
  }

  #[napi]
  pub fn from_f64_values(values: Vec<f64>) -> DynWinRTArray {
    let wvals: Vec<dynwinrt::WinRTValue> = values.into_iter().map(dynwinrt::WinRTValue::F64).collect();
    DynWinRTArray(dynwinrt::ArrayData::from_values(TABLE.f64_type(), &wvals))
  }

  #[napi]
  pub fn from_u8_values(values: Vec<u8>) -> DynWinRTArray {
    let wvals: Vec<dynwinrt::WinRTValue> = values.into_iter().map(dynwinrt::WinRTValue::U8).collect();
    DynWinRTArray(dynwinrt::ArrayData::from_values(TABLE.u8_type(), &wvals))
  }

  /// Wrap as DynWinRTValue::Array for passing to call().
  #[napi]
  pub fn to_value(&self) -> DynWinRTValue {
    DynWinRTValue(dynwinrt::WinRTValue::Array(self.0.clone()))
  }
}

// ======================================================================
// Struct binding — typed field access by index
// ======================================================================

#[napi]
pub struct DynWinRTStruct(dynwinrt::ValueTypeData);
unsafe impl Send for DynWinRTStruct {}
unsafe impl Sync for DynWinRTStruct {}

#[napi]
impl DynWinRTStruct {
  /// Create a zero-initialized struct of the given type.
  #[napi]
  pub fn create(typ: &DynWinRTType) -> DynWinRTStruct {
    DynWinRTStruct(typ.0.default_value())
  }

  #[napi]
  pub fn get_i32(&self, index: u32) -> i32 {
    self.0.get_field::<i32>(index as usize)
  }
  #[napi]
  pub fn set_i32(&mut self, index: u32, value: i32) {
    self.0.set_field(index as usize, value);
  }

  #[napi]
  pub fn get_u32(&self, index: u32) -> u32 {
    self.0.get_field::<u32>(index as usize)
  }
  #[napi]
  pub fn set_u32(&mut self, index: u32, value: u32) {
    self.0.set_field(index as usize, value);
  }

  #[napi]
  pub fn get_f32(&self, index: u32) -> f64 {
    self.0.get_field::<f32>(index as usize) as f64
  }
  #[napi]
  pub fn set_f32(&mut self, index: u32, value: f64) {
    self.0.set_field(index as usize, value as f32);
  }

  #[napi]
  pub fn get_f64(&self, index: u32) -> f64 {
    self.0.get_field::<f64>(index as usize)
  }
  #[napi]
  pub fn set_f64(&mut self, index: u32, value: f64) {
    self.0.set_field(index as usize, value);
  }

  #[napi]
  pub fn get_i64(&self, index: u32) -> i64 {
    self.0.get_field::<i64>(index as usize)
  }
  #[napi]
  pub fn set_i64(&mut self, index: u32, value: i64) {
    self.0.set_field(index as usize, value);
  }

  #[napi]
  pub fn get_u8(&self, index: u32) -> u32 {
    self.0.get_field::<u8>(index as usize) as u32
  }
  #[napi]
  pub fn set_u8(&mut self, index: u32, value: u32) {
    self.0.set_field(index as usize, value as u8);
  }

  /// Wrap as DynWinRTValue::Struct for passing to call().
  #[napi]
  pub fn to_value(&self) -> DynWinRTValue {
    DynWinRTValue(dynwinrt::WinRTValue::Struct(self.0.clone()))
  }
}

// ======================================================================
// IID constants
// ======================================================================

#[napi]
pub struct WinIIds(dynwinrt::IIds);

#[napi]
impl WinIIds {
  #[napi]
  pub fn IFileOpenPickerFactoryIID() -> WinGUID {
    WinGUID(dynwinrt::IIds::IFileOpenPickerFactory)
  }

  #[napi]
  pub fn ITextRecognizerStaticsIID() -> WinGUID {
    WinGUID(dynwinrt::IIds::ITextRecognizerStatics)
  }

  #[napi]
  pub fn IImageBufferStaticsIID() -> WinGUID {
    WinGUID(dynwinrt::IIds::IImageBufferStatics)
  }

  #[napi]
  pub fn ISoftwareBitmapIID() -> WinGUID {
    WinGUID(dynwinrt::IIds::ISoftwareBitmap)
  }

  #[napi]
  pub fn TextRecognizerIID() -> WinGUID {
    WinGUID(dynwinrt::IIds::TextRecognizer)
  }

  #[napi]
  pub fn RecognizedTextIID() -> WinGUID {
    WinGUID(dynwinrt::IIds::RecognizedText)
  }
}

// ======================================================================
// HTTP helpers
// ======================================================================

#[napi]
pub struct ComObj(IUnknown);
unsafe impl Send for ComObj {}
unsafe impl Sync for ComObj {}

impl Clone for ComObj {
  fn clone(&self) -> Self {
    ComObj(self.0.clone())
  }
}

#[napi]
pub fn http_client_get_sync(uri: String) -> Result<ComObj> {
  use windows::Foundation::Uri;
  use windows::Web::Http::HttpClient;
  let uri = Uri::CreateUri(&HSTRING::from(uri))
    .map_err(|e| Error::new(Status::GenericFailure, e.to_string()))?;
  let http_client =
    HttpClient::new().map_err(|e| Error::new(Status::GenericFailure, e.to_string()))?;
  let op = http_client
    .GetStringAsync(&uri)
    .map_err(|e| Error::new(Status::GenericFailure, e.to_string()))?;
  let r = op.as_raw();
  let ukn = unsafe { IUnknown::from_raw_borrowed(&r) }.unwrap();
  Ok(ComObj(ukn.clone()))
}

#[napi]
pub async fn async_progress_hstring_to_promise_string(obj: &ComObj) -> Result<String> {
  use windows::Web::Http::HttpProgress;
  let async_op: IAsyncOperationWithProgress<HSTRING, HttpProgress> = obj
    .0
    .cast()
    .map_err(|e| Error::new(Status::GenericFailure, e.to_string()))?;

  let r = async_op
    .await
    .map_err(|e| Error::new(Status::GenericFailure, e.to_string()))?;

  Ok(r.to_string())
}

// ======================================================================
// System info
// ======================================================================

#[napi]
pub fn has_package_identity() -> bool {
  use windows::ApplicationModel::AppInfo;
  match AppInfo::Current() {
    Ok(_) => true,
    Err(_) => false,
  }
}

#[napi]
pub fn get_computer_name() -> napi::Result<String> {
  #[cfg(target_os = "windows")]
  {
    use windows::core::PWSTR;
    use windows::Win32::System::WindowsProgramming::GetComputerNameW;

    let mut buffer = [0u16; 256];
    let mut size = buffer.len() as u32;

    unsafe {
      if GetComputerNameW(Some(PWSTR(buffer.as_mut_ptr())), &mut size).is_ok() {
        let name = String::from_utf16_lossy(&buffer[..size as usize]);
        Ok(name)
      } else {
        Err(napi::Error::from_reason("Failed to get computer name"))
      }
    }
  }

  #[cfg(not(target_os = "windows"))]
  {
    Err(napi::Error::from_reason(
      "This function is only available on Windows",
    ))
  }
}

#[napi]
pub fn get_windows_directory() -> napi::Result<String> {
  #[cfg(target_os = "windows")]
  {
    use windows::Win32::System::SystemInformation::GetWindowsDirectoryW;

    let mut buffer = [0u16; 260]; // MAX_PATH

    unsafe {
      let len = GetWindowsDirectoryW(Some(&mut buffer));
      if len > 0 {
        let path = String::from_utf16_lossy(&buffer[..len as usize]);
        Ok(path)
      } else {
        Err(napi::Error::from_reason("Failed to get Windows directory"))
      }
    }
  }

  #[cfg(not(target_os = "windows"))]
  {
    Err(napi::Error::from_reason(
      "This function is only available on Windows",
    ))
  }
}

// ======================================================================
// Demo — OCR
// ======================================================================

#[napi]
struct OCRDemo;

#[napi]
impl OCRDemo {
  #[napi]
  pub fn pick_image() -> DynWinRTValue {
    let bitmap = pollster::block_on(dynwinrt::get_bitmap_from_file());
    DynWinRTValue(bitmap)
  }

  #[napi]
  pub fn print_ocr_lines(text: &DynWinRTValue) {
    dynwinrt::print_ocr_paths(text.0.clone());
  }

  #[napi]
  pub async fn create_recognizer() -> DynWinRTValue {
    use dynwinrt::{IIds, WinRTValue};
    let factory =
      WinRTValue::from_activation_factory(h!("Microsoft.Windows.AI.Imaging.TextRecognizer"))
        .unwrap();
    let text_recongizer_factory = factory.cast(&IIds::ITextRecognizerStatics).unwrap();
    let ready_state = {
      let method = dynwinrt::MethodSignature::new(&*TABLE).add_out(TABLE.i32_type()).build(6);
      method.call_dynamic(text_recongizer_factory.as_object().unwrap().as_raw(), &[])
        .unwrap()[0].as_i32().unwrap()
    };
    println!("TextRecognizer ready state: {:?}", ready_state);

    if ready_state != 0 {
      panic!("TextRecognizer is not ready");
    }
    println!("Creating TextRecognizer asynchronously...");

    let recognizer_v = {
      let async_op_type = TABLE.async_operation(
        &TABLE.runtime_class(
          "Microsoft.Windows.AI.Imaging.TextRecognizer".into(),
          dynwinrt::IIds::TextRecognizer,
        ),
      );
      let create_async = dynwinrt::MethodSignature::new(&*TABLE).add_out(async_op_type).build(8);
      create_async
        .call_dynamic(text_recongizer_factory.as_object().unwrap().as_raw(), &[])
        .unwrap().into_iter().next().unwrap()
    };
    println!("TextRecognizer created successfully");
    let recognizer = (&recognizer_v).await.unwrap();
    DynWinRTValue(recognizer)
  }

  #[napi]
  pub fn create_image_buffer(bitmap: &DynWinRTValue) -> DynWinRTValue {
    use dynwinrt::{IIds, WinRTValue};
    let image_buffer_af =
      WinRTValue::from_activation_factory(h!("Microsoft.Graphics.Imaging.ImageBuffer")).unwrap();
    let image_buffer_static = image_buffer_af.cast(&IIds::IImageBufferStatics).unwrap();
    let create_for_bitmap = dynwinrt::MethodSignature::new(&*TABLE)
      .add_in(TABLE.object()).add_out(TABLE.object()).build(7);
    let image_buffer = create_for_bitmap
      .call_dynamic(image_buffer_static.as_object().unwrap().as_raw(),
        &[bitmap.0.cast(&IIds::ISoftwareBitmap).unwrap()])
      .unwrap().into_iter().next().unwrap()
      .as_object()
      .unwrap();
    println!("ImageBuffer created successfully");
    DynWinRTValue(WinRTValue::Object(image_buffer))
  }

  #[napi]
  pub async fn run_ocr_demo(
    recognizer: &DynWinRTValue,
    image_buffer: &DynWinRTValue,
  ) -> DynWinRTValue {
    use dynwinrt::IIds;
    let recognizer = &recognizer.0;
    let image_buffer = &image_buffer.0;
    let recognize = dynwinrt::MethodSignature::new(&*TABLE)
      .add_in(TABLE.object()).add_out(TABLE.object()).build(7);
    let res = recognize
      .call_dynamic(recognizer.as_object().unwrap().as_raw(), &[image_buffer.clone()])
      .unwrap().into_iter().next().unwrap();
    let result = res.cast(&IIds::RecognizedText).unwrap();
    println!("Text recognition completed successfully");
    DynWinRTValue(result)
  }
}

// ======================================================================
// Test helpers — used by JS test suite
// ======================================================================

#[napi]
pub fn use_dynwinrt_add(a: f64, b: f64) -> napi::Result<f64> {
  let result = dynwinrt::export_add(a, &b);
  Ok(result)
}

#[napi]
struct ComUri(windows::Foundation::Uri);

#[napi]
impl ComUri {
  #[napi]
  pub fn createTestUri(uri_str: String) -> napi::Result<ComUri> {
    #[cfg(target_os = "windows")]
    {
      use windows::Foundation::Uri;
      let hstring = HSTRING::from(uri_str);
      let uri = Uri::CreateUri(&hstring)
        .map_err(|e| napi::Error::from_reason(format!("Failed to create URI: {:?}", e)))?;
      Ok(ComUri(uri))
    }

    #[cfg(not(target_os = "windows"))]
    {
      Err(napi::Error::from_reason(
        "This function is only available on Windows",
      ))
    }
  }
}

#[napi]
struct DynWinRTVTable(dynwinrt::InterfaceSignature);

#[napi]
pub fn get_uri_vtable() -> DynWinRTVTable {
  DynWinRTVTable(dynwinrt::uri_vtable(&*TABLE))
}

#[napi]
pub fn callMethod(vtable: &DynWinRTVTable, index: i32, obj: &ComUri) -> String {
  let m = &vtable.0.methods[index as usize];
  let result = m.call_dynamic(obj.0.as_raw(), &[]).unwrap();
  result[0].as_hstring().unwrap().to_string()
}
