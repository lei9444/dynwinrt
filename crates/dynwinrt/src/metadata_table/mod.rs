mod type_kind;
mod type_handle;
mod value_data;
mod method_handle;
mod entries;
mod layout;
mod iid;

pub use type_kind::*;
pub use type_handle::TypeHandle;
pub use value_data::ValueTypeData;
pub use method_handle::MethodHandle;
pub use entries::*;

use std::collections::HashMap;
use std::sync::{Arc, RwLock};

use windows_core::GUID;

use crate::signature::{Method, MethodSignature};
use crate::value::WinRTValue;

use layout::{RuntimeClassData, ParameterizedData, StructEntry};

// ===========================================================================
// MetadataTable
// ===========================================================================

/// Centralized registry of WinRT types and methods.
///
/// Thread-safe, append-only, Arc-wrapped for shared ownership.
/// Stores type arenas (structs, parameterized types, etc.),
/// a global methods arena with pre-built Cif objects,
/// and named entries for interfaces, structs, enums, etc.
pub struct MetadataTable {
    // --- Type arenas ---
    structs: RwLock<Vec<StructEntry>>,
    runtime_classes: RwLock<Vec<RuntimeClassData>>,
    parameterized_types: RwLock<Vec<ParameterizedData>>,
    inner_types: RwLock<Vec<TypeKind>>,
    inner_type_pairs: RwLock<Vec<(TypeKind, TypeKind)>>,
    enum_names: RwLock<Vec<String>>,

    // --- Methods arena ---
    methods: RwLock<Vec<Method>>,

    // --- Named entries ---
    entries: RwLock<HashMap<String, TypeEntry>>,
}

impl std::fmt::Debug for MetadataTable {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("MetadataTable").finish_non_exhaustive()
    }
}

// Safety: MetadataTable is protected by RwLock internally.
// The non-Send/Sync raw pointers come from libffi Cif objects inside Method,
// which are only accessed through &self methods behind the RwLock.
unsafe impl Send for MetadataTable {}
unsafe impl Sync for MetadataTable {}

impl MetadataTable {
    pub fn new() -> Arc<Self> {
        Arc::new(MetadataTable {
            structs: RwLock::new(Vec::new()),
            runtime_classes: RwLock::new(Vec::new()),
            parameterized_types: RwLock::new(Vec::new()),
            inner_types: RwLock::new(Vec::new()),
            inner_type_pairs: RwLock::new(Vec::new()),
            enum_names: RwLock::new(Vec::new()),
            methods: RwLock::new(Vec::new()),
            entries: RwLock::new(HashMap::new()),
        })
    }

    // -----------------------------------------------------------------------
    // Type factory methods
    // -----------------------------------------------------------------------

    pub fn make(self: &Arc<Self>, kind: TypeKind) -> TypeHandle {
        TypeHandle {
            table: Arc::clone(self),
            kind,
        }
    }

    // Primitive types
    pub fn bool_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::Bool) }
    pub fn i8_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::I8) }
    pub fn u8_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::U8) }
    pub fn i16_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::I16) }
    pub fn u16_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::U16) }
    pub fn char16_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::Char16) }
    pub fn i32_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::I32) }
    pub fn u32_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::U32) }
    pub fn i64_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::I64) }
    pub fn u64_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::U64) }
    pub fn f32_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::F32) }
    pub fn f64_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::F64) }
    pub fn guid_type(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::Guid) }
    pub fn hstring(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::HString) }
    pub fn object(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::Object) }
    pub fn hresult(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::HResult) }
    pub fn array_of_iunknown(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::ArrayOfIUnknown) }
    pub fn async_action(self: &Arc<Self>) -> TypeHandle { self.make(TypeKind::IAsyncAction) }

    /// Create a TypeHandle from a TypeKind. Only works for simple (non-indexed) kinds.
    pub fn handle_from_kind(self: &Arc<Self>, kind: TypeKind) -> TypeHandle {
        self.make(kind)
    }

    // GUID-carrying types
    pub fn interface(self: &Arc<Self>, iid: GUID) -> TypeHandle {
        self.make(TypeKind::Interface(iid))
    }
    pub fn delegate(self: &Arc<Self>, iid: GUID) -> TypeHandle {
        self.make(TypeKind::Delegate(iid))
    }
    pub fn generic(self: &Arc<Self>, piid: GUID, arity: u32) -> TypeHandle {
        self.make(TypeKind::Generic { piid, arity })
    }

    // Compound types that allocate indexed storage
    pub fn runtime_class(self: &Arc<Self>, name: String, default_iid: GUID) -> TypeHandle {
        let mut rcs = self.runtime_classes.write().unwrap();
        let idx = rcs.len() as u32;
        rcs.push(RuntimeClassData { name, default_iid });
        self.make(TypeKind::RuntimeClass(idx))
    }

    pub fn parameterized(self: &Arc<Self>, generic_def: &TypeHandle, args: &[TypeHandle]) -> TypeHandle {
        let mut pts = self.parameterized_types.write().unwrap();
        let idx = pts.len() as u32;
        pts.push(ParameterizedData {
            generic_def: generic_def.kind,
            args: args.iter().map(|a| a.kind).collect(),
        });
        self.make(TypeKind::Parameterized(idx))
    }

    pub fn async_operation(self: &Arc<Self>, result_type: &TypeHandle) -> TypeHandle {
        let idx = self.register_inner_type(result_type.kind);
        self.make(TypeKind::IAsyncOperation(idx))
    }

    pub fn async_action_with_progress(self: &Arc<Self>, progress_type: &TypeHandle) -> TypeHandle {
        let idx = self.register_inner_type(progress_type.kind);
        self.make(TypeKind::IAsyncActionWithProgress(idx))
    }

    pub fn async_operation_with_progress(
        self: &Arc<Self>,
        result_type: &TypeHandle,
        progress_type: &TypeHandle,
    ) -> TypeHandle {
        let idx = self.register_inner_type_pair(result_type.kind, progress_type.kind);
        self.make(TypeKind::IAsyncOperationWithProgress(idx))
    }

    pub fn out_value(self: &Arc<Self>, inner: &TypeHandle) -> TypeHandle {
        let idx = self.register_inner_type(inner.kind);
        self.make(TypeKind::OutValue(idx))
    }

    pub fn array(self: &Arc<Self>, element_type: &TypeHandle) -> TypeHandle {
        let idx = self.register_inner_type(element_type.kind);
        self.make(TypeKind::Array(idx))
    }

    pub fn define_struct(self: &Arc<Self>, fields: &[TypeHandle]) -> TypeHandle {
        let field_kinds: Vec<TypeKind> = fields.iter().map(|h| h.kind).collect();
        let (field_offsets, layout) = self.compute_layout(&field_kinds);
        let mut structs = self.structs.write().unwrap();
        let id = structs.len() as u32;
        structs.push(StructEntry {
            name: None,
            field_kinds,
            field_offsets,
            layout,
        });
        self.make(TypeKind::Struct(id))
    }

    /// Define a named struct type with WinRT full name (for correct IID signature computation).
    pub fn define_named_struct(self: &Arc<Self>, winrt_name: &str, fields: &[TypeHandle]) -> TypeHandle {
        let field_kinds: Vec<TypeKind> = fields.iter().map(|h| h.kind).collect();
        let (field_offsets, layout) = self.compute_layout(&field_kinds);
        let mut structs = self.structs.write().unwrap();
        let id = structs.len() as u32;
        structs.push(StructEntry {
            name: Some(winrt_name.to_string()),
            field_kinds,
            field_offsets,
            layout,
        });
        self.make(TypeKind::Struct(id))
    }

    /// Define a named enum type (ABI = i32, but carries name for signature).
    pub fn define_named_enum(self: &Arc<Self>, winrt_name: &str) -> TypeHandle {
        let mut enums = self.enum_names.write().unwrap();
        let id = enums.len() as u32;
        enums.push(winrt_name.to_string());
        self.make(TypeKind::Enum(id))
    }

    // -----------------------------------------------------------------------
    // Methods arena
    // -----------------------------------------------------------------------

    pub(crate) fn invoke_method(
        &self,
        index: u32,
        obj: *mut std::ffi::c_void,
        args: &[WinRTValue],
    ) -> windows_core::Result<Vec<WinRTValue>> {
        let methods = self.methods.read().unwrap();
        methods[index as usize].call_dynamic(obj, args)
    }

    pub(crate) fn add_method_to_interface(&self, iid: &GUID, name: &str, sig: MethodSignature) -> u32 {
        let mut entries = self.entries.write().unwrap();
        let entry = entries.values_mut()
            .find(|e| matches!(e, TypeEntry::Interface(ie) if ie.iid == *iid))
            .expect("Interface not found for add_method");

        if let TypeEntry::Interface(ie) = entry {
            let vtable_index = 6 + ie.method_indices.len();
            let method = sig.build(vtable_index);
            let arena_index = {
                let mut methods = self.methods.write().unwrap();
                let idx = methods.len() as u32;
                methods.push(method);
                idx
            };
            ie.method_names.push(name.to_string());
            ie.method_indices.push(arena_index);
            vtable_index as u32
        } else {
            unreachable!()
        }
    }

    // -----------------------------------------------------------------------
    // Registration API
    // -----------------------------------------------------------------------

    pub fn register_interface(self: &Arc<Self>, name: &str, iid: GUID) -> TypeHandle {
        let entry = TypeEntry::Interface(InterfaceEntry {
            name: name.to_string(),
            iid,
            method_names: Vec::new(),
            method_indices: Vec::new(),
        });
        self.entries.write().unwrap().insert(name.to_string(), entry);
        self.make(TypeKind::Interface(iid))
    }

    pub fn register_runtime_class(
        &self,
        name: &str,
        default_iid: GUID,
        interfaces: Vec<InterfaceRef>,
        has_default_constructor: bool,
    ) {
        let entry = TypeEntry::RuntimeClass(RuntimeClassEntry {
            name: name.to_string(),
            default_iid,
            interfaces,
            has_default_constructor,
        });
        self.entries.write().unwrap().insert(name.to_string(), entry);
    }

    pub fn register_struct(
        self: &Arc<Self>,
        name: &str,
        fields: Vec<(String, TypeHandle)>,
    ) -> TypeHandle {
        {
            let entries = self.entries.read().unwrap();
            if let Some(TypeEntry::Struct(s)) = entries.get(name) {
                return s.type_handle.clone();
            }
        }

        let (field_names, field_types): (Vec<String>, Vec<TypeHandle>) =
            fields.into_iter().unzip();

        let type_handle = self.define_struct(&field_types);

        let entry = TypeEntry::Struct(StructTypeEntry {
            name: name.to_string(),
            type_handle: type_handle.clone(),
            field_names,
            field_types,
        });
        self.entries.write().unwrap().insert(name.to_string(), entry);
        type_handle
    }

    pub fn register_enum(
        &self,
        name: &str,
        underlying: TypeHandle,
        members: Vec<(String, i32)>,
    ) {
        let entry = TypeEntry::Enum(EnumEntry {
            name: name.to_string(),
            underlying,
            members,
        });
        self.entries.write().unwrap().insert(name.to_string(), entry);
    }

    pub fn register_delegate(&self, name: &str, iid: GUID) {
        let entry = TypeEntry::Delegate(DelegateEntry {
            name: name.to_string(),
            iid,
        });
        self.entries.write().unwrap().insert(name.to_string(), entry);
    }

    // -----------------------------------------------------------------------
    // Query API
    // -----------------------------------------------------------------------

    pub fn contains(&self, name: &str) -> bool {
        self.entries.read().unwrap().contains_key(name)
    }

    pub fn get_iid(&self, name: &str) -> Option<GUID> {
        let entries = self.entries.read().unwrap();
        match entries.get(name)? {
            TypeEntry::Interface(e) => Some(e.iid),
            TypeEntry::Delegate(e) => Some(e.iid),
            TypeEntry::RuntimeClass(e) => Some(e.default_iid),
            _ => None,
        }
    }

    pub fn get_struct_handle(&self, name: &str) -> Option<TypeHandle> {
        let entries = self.entries.read().unwrap();
        match entries.get(name)? {
            TypeEntry::Struct(e) => Some(e.type_handle.clone()),
            _ => None,
        }
    }

    pub fn get_type_handle(self: &Arc<Self>, name: &str) -> Option<TypeHandle> {
        let entries = self.entries.read().unwrap();
        match entries.get(name)? {
            TypeEntry::Interface(e) => Some(self.interface(e.iid)),
            TypeEntry::RuntimeClass(e) => {
                Some(self.runtime_class(e.name.clone(), e.default_iid))
            }
            TypeEntry::Struct(e) => Some(e.type_handle.clone()),
            TypeEntry::Enum(e) => Some(e.underlying.clone()),
            TypeEntry::Delegate(e) => Some(self.delegate(e.iid)),
        }
    }

    pub fn get_enum_value(&self, enum_name: &str, member_name: &str) -> Option<i32> {
        let entries = self.entries.read().unwrap();
        match entries.get(enum_name)? {
            TypeEntry::Enum(e) => e
                .members
                .iter()
                .find(|(n, _)| n == member_name)
                .map(|(_, v)| *v),
            _ => None,
        }
    }

    // -----------------------------------------------------------------------
    // Call API (uses cached Cifs)
    // -----------------------------------------------------------------------

    pub fn call(
        &self,
        interface_name: &str,
        method_name: &str,
        obj: *mut std::ffi::c_void,
        args: &[WinRTValue],
    ) -> crate::result::Result<Vec<WinRTValue>> {
        let entries = self.entries.read().unwrap();
        let entry = entries
            .get(interface_name)
            .ok_or_else(|| crate::result::Error::TypeNotFound(interface_name.to_string()))?;
        let iface = match entry {
            TypeEntry::Interface(e) => e,
            _ => {
                return Err(crate::result::Error::NotAnInterface(
                    interface_name.to_string(),
                ))
            }
        };
        let pos = iface
            .method_names
            .iter()
            .position(|n| n == method_name)
            .ok_or_else(|| {
                crate::result::Error::MethodNotFound(
                    interface_name.to_string(),
                    method_name.to_string(),
                )
            })?;
        let arena_index = iface.method_indices[pos];
        drop(entries);
        self.invoke_method(arena_index, obj, args)
            .map_err(|e| crate::result::Error::WindowsError(e))
    }

    pub fn call_by_index(
        &self,
        interface_name: &str,
        method_index: usize,
        obj: *mut std::ffi::c_void,
        args: &[WinRTValue],
    ) -> crate::result::Result<Vec<WinRTValue>> {
        let entries = self.entries.read().unwrap();
        let entry = entries
            .get(interface_name)
            .ok_or_else(|| crate::result::Error::TypeNotFound(interface_name.to_string()))?;
        let iface = match entry {
            TypeEntry::Interface(e) => e,
            _ => {
                return Err(crate::result::Error::NotAnInterface(
                    interface_name.to_string(),
                ))
            }
        };
        let arena_index = *iface.method_indices.get(method_index).ok_or_else(|| {
            crate::result::Error::MethodNotFound(
                interface_name.to_string(),
                format!("[index {}]", method_index),
            )
        })?;
        drop(entries);
        self.invoke_method(arena_index, obj, args)
            .map_err(|e| crate::result::Error::WindowsError(e))
    }

    pub(crate) fn method_by_vtable_index(self: &Arc<Self>, iid: &GUID, vtable_index: usize) -> Option<MethodHandle> {
        if vtable_index < 6 {
            return None;
        }
        let local_index = vtable_index - 6;
        let entries = self.entries.read().unwrap();
        for entry in entries.values() {
            if let TypeEntry::Interface(ie) = entry {
                if ie.iid == *iid {
                    let arena_index = *ie.method_indices.get(local_index)?;
                    return Some(MethodHandle {
                        table: Arc::clone(self),
                        index: arena_index,
                    });
                }
            }
        }
        None
    }

    pub(crate) fn method_by_name(self: &Arc<Self>, iid: &GUID, name: &str) -> Option<MethodHandle> {
        let entries = self.entries.read().unwrap();
        for entry in entries.values() {
            if let TypeEntry::Interface(ie) = entry {
                if ie.iid == *iid {
                    let pos = ie.method_names.iter().position(|n| n == name)?;
                    let arena_index = ie.method_indices[pos];
                    return Some(MethodHandle {
                        table: Arc::clone(self),
                        index: arena_index,
                    });
                }
            }
        }
        None
    }
}

// ===========================================================================
// Tests
// ===========================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::abi::AbiType;
    use windows_core::Interface;

    #[test]
    fn primitive_sizes() {
        let table = MetadataTable::new();
        assert_eq!(table.f32_type().size_of(), 4);
        assert_eq!(table.f64_type().size_of(), 8);
        assert_eq!(table.i32_type().size_of(), 4);
        assert_eq!(table.u8_type().size_of(), 1);
        assert_eq!(table.f32_type().align_of(), 4);
        assert_eq!(table.f64_type().align_of(), 8);
    }

    #[test]
    fn point_layout() {
        let table = MetadataTable::new();
        let f32_h = table.f32_type();
        let point = table.define_struct(&[f32_h.clone(), f32_h]);
        assert_eq!(point.size_of(), 8);
        assert_eq!(point.align_of(), 4);
        assert_eq!(point.field_count(), 2);
        assert_eq!(point.field_offset(0), 0);
        assert_eq!(point.field_offset(1), 4);
    }

    #[test]
    fn rect_layout() {
        let table = MetadataTable::new();
        let f32_h = table.f32_type();
        let rect = table.define_struct(&[f32_h.clone(), f32_h.clone(), f32_h.clone(), f32_h]);
        assert_eq!(rect.size_of(), 16);
        assert_eq!(rect.align_of(), 4);
    }

    #[test]
    fn basic_geoposition_layout() {
        let table = MetadataTable::new();
        let f64_h = table.f64_type();
        let geo = table.define_struct(&[f64_h.clone(), f64_h.clone(), f64_h]);
        assert_eq!(geo.size_of(), 24);
        assert_eq!(geo.align_of(), 8);
    }

    #[test]
    fn value_type_data_field_access() {
        let table = MetadataTable::new();
        let f32_h = table.f32_type();
        let point = table.define_struct(&[f32_h.clone(), f32_h]);
        let mut val = point.default_value();
        val.set_field(0, 10.0f32);
        val.set_field(1, 20.0f32);
        assert_eq!(val.get_field::<f32>(0), 10.0);
        assert_eq!(val.get_field::<f32>(1), 20.0);
    }

    #[test]
    fn value_type_matches_windows_point_layout() {
        use windows::Foundation::Point;
        let table = MetadataTable::new();
        let f32_h = table.f32_type();
        let point = table.define_struct(&[f32_h.clone(), f32_h]);
        assert_eq!(point.size_of(), std::mem::size_of::<Point>());
        assert_eq!(point.align_of(), std::mem::align_of::<Point>());
    }

    #[test]
    fn libffi_type_struct() {
        let table = MetadataTable::new();
        let f32_h = table.f32_type();
        let f64_h = table.f64_type();
        let point = table.define_struct(&[f32_h.clone(), f32_h]);
        let _ = point.libffi_type();
        let outer = table.define_struct(&[point, f64_h]);
        let _ = outer.libffi_type();
    }

    #[test]
    fn geopoint_create_via_table() -> windows::core::Result<()> {
        use libffi::middle::{Cif, CodePtr, arg};
        use windows::Devices::Geolocation::Geopoint;
        use windows::Win32::System::WinRT::{
            IActivationFactory, RO_INIT_MULTITHREADED, RoGetActivationFactory, RoInitialize,
        };
        use windows::core::{Interface, h};
        use windows_core::HRESULT;
        use crate::call::get_vtable_function_ptr;

        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };
        let table = MetadataTable::new();
        let f64_h = table.f64_type();
        let geo_type = table.define_struct(&[f64_h.clone(), f64_h.clone(), f64_h]);
        let mut geo_val = geo_type.default_value();
        geo_val.set_field(0, 47.643f64);
        geo_val.set_field(1, -122.131f64);
        geo_val.set_field(2, 100.0f64);

        let afactory = unsafe {
            RoGetActivationFactory::<IActivationFactory>(h!("Windows.Devices.Geolocation.Geopoint"))
        }?;
        let geopoint_factory = afactory.cast::<windows::Devices::Geolocation::IGeopointFactory>()?;
        let fptr = get_vtable_function_ptr(geopoint_factory.as_raw(), 6);
        let cif = Cif::new(
            vec![libffi::middle::Type::pointer(), geo_type.libffi_type(), libffi::middle::Type::pointer()],
            libffi::middle::Type::i32(),
        );
        let mut out: *mut std::ffi::c_void = std::ptr::null_mut();
        let this = geopoint_factory.as_raw();
        let geo_data_ref = unsafe { &*geo_val.as_ptr() };
        let hr: HRESULT = unsafe { cif.call(CodePtr(fptr), &[arg(&this), arg(geo_data_ref), arg(&(&mut out))]) };
        hr.ok()?;
        let geopoint = unsafe { Geopoint::from_raw(out) };
        let pos = geopoint.Position()?;
        assert!((pos.Latitude - 47.643).abs() < 1e-6);
        assert!((pos.Longitude - (-122.131)).abs() < 1e-6);
        Ok(())
    }

    #[test]
    fn register_and_query_struct() {
        let table = MetadataTable::new();
        let handle = table.register_struct(
            "Windows.Foundation.Point",
            vec![("X".into(), table.f32_type()), ("Y".into(), table.f32_type())],
        );
        assert_eq!(handle.size_of(), 8);
        assert!(table.contains("Windows.Foundation.Point"));
    }

    #[test]
    fn register_and_query_interface() {
        let iid = GUID::from_u128(0x9E365E57_48B2_4160_956F_C7385120BBFC);
        let table = MetadataTable::new();
        table.register_interface("Windows.Foundation.IUriRuntimeClass", iid)
            .add_method("get_AbsoluteUri", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_DisplayUri", MethodSignature::new(&table).add_out(table.hstring()));
        assert_eq!(table.get_iid("Windows.Foundation.IUriRuntimeClass"), Some(iid));
        assert!(table.contains("Windows.Foundation.IUriRuntimeClass"));
    }

    #[test]
    fn method_handle_by_vtable_index() {
        let iid = GUID::from_u128(0x9E365E57_48B2_4160_956F_C7385120BBFC);
        let table = MetadataTable::new();
        let iface = table.register_interface("IUri", iid)
            .add_method("get_AbsoluteUri", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_DisplayUri", MethodSignature::new(&table).add_out(table.hstring()));
        assert!(iface.method(6).is_some());
        assert!(iface.method(7).is_some());
        assert!(iface.method(8).is_none());
        assert!(iface.method(5).is_none());
        assert!(iface.method_by_name("get_AbsoluteUri").is_some());
        assert!(iface.method_by_name("nonexistent").is_none());
    }

    #[test]
    fn call_uri_via_table() {
        use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};
        use windows_core::{Interface, h};

        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };
        let table = MetadataTable::new();

        let factory_iid = GUID::from_u128(0x44A9796F_723E_4FDF_A218_033E75B0C084);
        table.register_interface("Windows.Foundation.IUriRuntimeClassFactory", factory_iid)
            .add_method("CreateUri", MethodSignature::new(&table).add_in(table.hstring()).add_out(table.object()));

        let uri_iid = GUID::from_u128(0x9E365E57_48B2_4160_956F_C7385120BBFC);
        table.register_interface("Windows.Foundation.IUriRuntimeClass", uri_iid)
            .add_method("get_AbsoluteUri", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_DisplayUri", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_Domain", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_Extension", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_Fragment", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_Host", MethodSignature::new(&table).add_out(table.hstring()));

        let factory = unsafe { windows::Win32::System::WinRT::RoGetActivationFactory::<windows::Win32::System::WinRT::IActivationFactory>(h!("Windows.Foundation.Uri")) }.unwrap();
        let uri_factory: windows_core::IUnknown = factory.cast().unwrap();
        let mut factory_ptr = std::ptr::null_mut();
        unsafe { uri_factory.query(&factory_iid, &mut factory_ptr).ok().unwrap(); }

        let results = table.call("Windows.Foundation.IUriRuntimeClassFactory", "CreateUri", factory_ptr, &[WinRTValue::HString(windows_core::HSTRING::from("https://www.example.com/path?query=1#fragment"))]).unwrap();
        let uri_obj = results[0].as_object().unwrap();
        let mut uri_ptr = std::ptr::null_mut();
        unsafe { uri_obj.query(&uri_iid, &mut uri_ptr).ok().unwrap(); }

        let host = table.call("Windows.Foundation.IUriRuntimeClass", "get_Host", uri_ptr, &[]).unwrap()[0].as_hstring().unwrap();
        assert_eq!(host.to_string(), "www.example.com");

        let abs_uri = table.call_by_index("Windows.Foundation.IUriRuntimeClass", 0, uri_ptr, &[]).unwrap()[0].as_hstring().unwrap();
        assert_eq!(abs_uri.to_string(), "https://www.example.com/path?query=1#fragment");
    }

    #[test]
    fn test_abi_type_mapping() {
        let table = MetadataTable::new();
        assert_eq!(table.bool_type().abi_type(), AbiType::Bool);
        assert_eq!(table.i32_type().abi_type(), AbiType::I32);
        assert_eq!(table.hstring().abi_type(), AbiType::Ptr);
    }

    #[test]
    fn test_signature_string() {
        let table = MetadataTable::new();
        assert_eq!(table.i32_type().signature_string(), "i4");
        assert_eq!(table.hstring().signature_string(), "string");
        let g = table.generic(IASYNC_OPERATION, 1);
        let sig = table.parameterized(&g, &[table.hstring()]);
        assert_eq!(sig.signature_string(), "pinterface({9fc2b0bb-e446-44e2-aa61-9cab8f636af2};string)");
    }

    #[test]
    fn test_iid() {
        let table = MetadataTable::new();
        let iid = GUID::from_u128(0x12345678_1234_1234_1234_123456789abc);
        assert_eq!(table.interface(iid).iid(), Some(iid));
        let g = table.generic(IASYNC_OPERATION, 1);
        let p = table.parameterized(&g, &[table.hstring()]);
        assert_eq!(p.iid().unwrap(), windows_future::IAsyncOperation::<windows_core::HSTRING>::IID);
    }

    #[test]
    fn test_iid_runtime_class_as_type_arg() {
        let table = MetadataTable::new();
        let storage_file = table.runtime_class("Windows.Storage.StorageFile".into(), GUID::from_u128(0xFA3F6186_4214_428C_A64C_14C9AC7315EA));
        let g = table.generic(IASYNC_OPERATION, 1);
        let ty = table.parameterized(&g, &[storage_file]);
        assert_eq!(ty.iid().unwrap(), windows_future::IAsyncOperation::<windows::Storage::StorageFile>::IID);
    }

    #[test]
    fn test_guid_braced_format() {
        let guid = GUID::from_u128(0x9fc2b0bb_e446_44e2_aa61_9cab8f636af2);
        assert_eq!(format_guid_braced(&guid), "{9fc2b0bb-e446-44e2-aa61-9cab8f636af2}");
    }
}
