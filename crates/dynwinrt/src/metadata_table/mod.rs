mod type_kind;
mod type_handle;
mod value_data;
mod method_handle;
mod arena;
mod iid;

pub use type_kind::*;
pub use type_handle::TypeHandle;
pub use value_data::ValueTypeData;
pub use method_handle::MethodHandle;

use std::collections::HashMap;
use std::sync::{Arc, RwLock};

use windows_core::GUID;

use crate::signature::{Method, MethodSignature};

use arena::*;

// ===========================================================================
// MetadataTable
// ===========================================================================

/// Centralized registry of WinRT types and methods.
///
/// Thread-safe, append-only, Arc-wrapped for shared ownership.
/// All data lives in arenas; lightweight indexes provide O(1) lookup.
/// All direct arena access goes through `arena.rs` methods.
pub struct MetadataTable {
    // --- Type arenas (primary data) ---
    structs: RwLock<Vec<StructEntry>>,
    runtime_classes: RwLock<Vec<RuntimeClassData>>,
    parameterized_types: RwLock<Vec<ParameterizedData>>,
    inner_types: RwLock<Vec<TypeKind>>,
    inner_type_pairs: RwLock<Vec<(TypeKind, TypeKind)>>,
    enum_entries: RwLock<Vec<EnumData>>,

    // --- Methods arena ---
    methods: RwLock<Vec<Method>>,

    // --- Indexes (no data duplication, only pointers) ---
    /// IID → method table for O(1) interface method lookup.
    interface_methods: RwLock<HashMap<GUID, InterfaceMethodTable>>,
    /// Name → struct arena index for dedup on struct_type.
    struct_names: RwLock<HashMap<String, u32>>,
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
            enum_entries: RwLock::new(Vec::new()),
            methods: RwLock::new(Vec::new()),
            interface_methods: RwLock::new(HashMap::new()),
            struct_names: RwLock::new(HashMap::new()),
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
        self.make(self.push_runtime_class(name, default_iid))
    }

    pub fn parameterized(self: &Arc<Self>, generic_def: &TypeHandle, args: &[TypeHandle]) -> TypeHandle {
        let args_kinds: Vec<TypeKind> = args.iter().map(|a| a.kind).collect();
        self.make(self.push_parameterized(generic_def.kind, args_kinds))
    }

    pub fn async_operation(self: &Arc<Self>, result_type: &TypeHandle) -> TypeHandle {
        let idx = self.push_inner_type(result_type.kind);
        self.make(TypeKind::IAsyncOperation(idx))
    }

    pub fn async_action_with_progress(self: &Arc<Self>, progress_type: &TypeHandle) -> TypeHandle {
        let idx = self.push_inner_type(progress_type.kind);
        self.make(TypeKind::IAsyncActionWithProgress(idx))
    }

    pub fn async_operation_with_progress(
        self: &Arc<Self>,
        result_type: &TypeHandle,
        progress_type: &TypeHandle,
    ) -> TypeHandle {
        let idx = self.push_inner_type_pair(result_type.kind, progress_type.kind);
        self.make(TypeKind::IAsyncOperationWithProgress(idx))
    }

    pub fn out_value(self: &Arc<Self>, inner: &TypeHandle) -> TypeHandle {
        let idx = self.push_inner_type(inner.kind);
        self.make(TypeKind::OutValue(idx))
    }

    pub fn array(self: &Arc<Self>, element_type: &TypeHandle) -> TypeHandle {
        let idx = self.push_inner_type(element_type.kind);
        self.make(TypeKind::Array(idx))
    }

    // -----------------------------------------------------------------------
    // Registration API (single entry point for each type)
    // -----------------------------------------------------------------------

    /// Register a named interface. Creates an IID → method table.
    /// Returns a TypeHandle for chaining `.add_method()`.
    pub fn register_interface(self: &Arc<Self>, _name: &str, iid: GUID) -> TypeHandle {
        self.create_interface_method_table(iid);
        self.make(TypeKind::Interface(iid))
    }

    /// Register a named struct with dedup. If already registered, returns
    /// the existing TypeHandle.
    pub fn struct_type(self: &Arc<Self>, name: &str, fields: &[TypeHandle]) -> TypeHandle {
        if let Some(idx) = self.get_struct_index_by_name(name) {
            return self.make(TypeKind::Struct(idx));
        }
        let field_kinds: Vec<TypeKind> = fields.iter().map(|h| h.kind).collect();
        let id = self.push_struct(name, field_kinds);
        self.insert_struct_name(name, id);
        self.make(TypeKind::Struct(id))
    }

    /// Register a named enum with member values.
    pub fn enum_type(self: &Arc<Self>, name: &str, members: Vec<(String, i32)>) -> TypeHandle {
        let id = self.push_enum(name, members);
        self.make(TypeKind::Enum(id))
    }

    // -----------------------------------------------------------------------
    // Methods
    // -----------------------------------------------------------------------

    /// Add a method to the interface identified by IID.
    pub(crate) fn add_method_to_interface(&self, iid: &GUID, name: &str, sig: MethodSignature) -> u32 {
        self.push_method(iid, name, sig)
    }

    /// Get a MethodHandle by vtable index. O(1) lookup by IID.
    pub(crate) fn method_by_vtable_index(self: &Arc<Self>, iid: &GUID, vtable_index: usize) -> Option<MethodHandle> {
        let arena_index = self.get_method_arena_index_by_vtable(iid, vtable_index)?;
        Some(MethodHandle { table: Arc::clone(self), index: arena_index })
    }

    /// Get a MethodHandle by method name. O(1) IID lookup + linear name scan.
    pub(crate) fn method_by_name(self: &Arc<Self>, iid: &GUID, name: &str) -> Option<MethodHandle> {
        let arena_index = self.get_method_arena_index_by_name(iid, name)?;
        Some(MethodHandle { table: Arc::clone(self), index: arena_index })
    }

    // -----------------------------------------------------------------------
    // Query API
    // -----------------------------------------------------------------------

    pub fn get_enum_value(&self, enum_name: &str, member_name: &str) -> Option<i32> {
        self.get_enum_members(enum_name)?
            .iter()
            .find(|(n, _)| n == member_name)
            .map(|(_, v)| *v)
    }
}

// ===========================================================================
// Tests
// ===========================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::abi::AbiType;
    use crate::value::WinRTValue;
    use windows_core::Interface;

    // -----------------------------------------------------------------------
    // Primitive types
    // -----------------------------------------------------------------------

    #[test]
    fn primitive_size_and_align() {
        let table = MetadataTable::new();
        assert_eq!(table.u8_type().size_of(), 1);
        assert_eq!(table.i32_type().size_of(), 4);
        assert_eq!(table.f32_type().size_of(), 4);
        assert_eq!(table.f64_type().size_of(), 8);
        assert_eq!(table.f32_type().align_of(), 4);
        assert_eq!(table.f64_type().align_of(), 8);
    }

    #[test]
    fn abi_type_mapping() {
        let table = MetadataTable::new();
        assert_eq!(table.bool_type().abi_type(), AbiType::Bool);
        assert_eq!(table.i32_type().abi_type(), AbiType::I32);
        assert_eq!(table.f64_type().abi_type(), AbiType::F64);
        assert_eq!(table.hstring().abi_type(), AbiType::Ptr);
        assert_eq!(table.object().abi_type(), AbiType::Ptr);
    }

    // -----------------------------------------------------------------------
    // Struct: layout, field access, libffi, Windows ABI compatibility
    // -----------------------------------------------------------------------

    #[test]
    fn struct_layout_and_field_access() {
        let table = MetadataTable::new();
        let f32_h = table.f32_type();
        let point = table.struct_type("Windows.Foundation.Point", &[f32_h.clone(), f32_h]);

        // Layout
        assert_eq!(point.size_of(), 8);
        assert_eq!(point.align_of(), 4);
        assert_eq!(point.field_count(), 2);
        assert_eq!(point.field_offset(0), 0);
        assert_eq!(point.field_offset(1), 4);

        // Matches real Windows.Foundation.Point
        assert_eq!(point.size_of(), std::mem::size_of::<windows::Foundation::Point>());
        assert_eq!(point.align_of(), std::mem::align_of::<windows::Foundation::Point>());

        // Field read/write
        let mut val = point.default_value();
        val.set_field(0, 10.0f32);
        val.set_field(1, 20.0f32);
        assert_eq!(val.get_field::<f32>(0), 10.0);
        assert_eq!(val.get_field::<f32>(1), 20.0);
    }

    #[test]
    fn struct_mixed_alignment() {
        // BasicGeoposition has f64 fields — tests 8-byte alignment
        let table = MetadataTable::new();
        let f64_h = table.f64_type();
        let geo = table.struct_type(
            "Windows.Devices.Geolocation.BasicGeoposition",
            &[f64_h.clone(), f64_h.clone(), f64_h],
        );
        assert_eq!(geo.size_of(), 24);
        assert_eq!(geo.align_of(), 8);
    }

    #[test]
    fn struct_nested_libffi_type() {
        let table = MetadataTable::new();
        let f32_h = table.f32_type();
        let f64_h = table.f64_type();
        let point = table.struct_type("Windows.Foundation.Point", &[f32_h.clone(), f32_h]);
        let _ = point.libffi_type(); // should not panic

        let outer = table.struct_type("Test.PointWithAltitude", &[point, f64_h]);
        let _ = outer.libffi_type(); // nested struct should work
    }

    #[test]
    fn struct_dedup_by_name() {
        let table = MetadataTable::new();
        let f32_h = table.f32_type();
        let h1 = table.struct_type("Windows.Foundation.Point", &[f32_h.clone(), f32_h.clone()]);
        let h2 = table.struct_type("Windows.Foundation.Point", &[f32_h.clone(), f32_h]);

        // Same TypeKind (same arena index)
        assert_eq!(h1.kind(), h2.kind());
        assert_eq!(h1.size_of(), h2.size_of());
    }

    // -----------------------------------------------------------------------
    // Enum
    // -----------------------------------------------------------------------

    #[test]
    fn enum_registration_and_query() {
        let table = MetadataTable::new();
        let handle = table.enum_type("Windows.Foundation.AsyncStatus", vec![
            ("Started".into(), 0),
            ("Completed".into(), 1),
            ("Canceled".into(), 2),
            ("Error".into(), 3),
        ]);

        // ABI is i32
        assert_eq!(handle.abi_type(), AbiType::I32);

        // Query by name
        assert_eq!(table.get_enum_value("Windows.Foundation.AsyncStatus", "Completed"), Some(1));
        assert_eq!(table.get_enum_value("Windows.Foundation.AsyncStatus", "Error"), Some(3));
        assert_eq!(table.get_enum_value("Windows.Foundation.AsyncStatus", "Nonexistent"), None);
        assert_eq!(table.get_enum_value("Nonexistent.Enum", "Foo"), None);
    }

    // -----------------------------------------------------------------------
    // Interface: registration, method lookup
    // -----------------------------------------------------------------------

    #[test]
    fn interface_method_lookup() {
        let iid = GUID::from_u128(0x9E365E57_48B2_4160_956F_C7385120BBFC);
        let table = MetadataTable::new();
        let iface = table.register_interface("IUriRuntimeClass", iid)
            .add_method("get_AbsoluteUri", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_DisplayUri", MethodSignature::new(&table).add_out(table.hstring()));

        // By vtable index (6 = first user method after IInspectable)
        assert!(iface.method(6).is_some());
        assert!(iface.method(7).is_some());
        assert!(iface.method(8).is_none()); // out of bounds
        assert!(iface.method(5).is_none()); // IInspectable range

        // By name
        assert!(iface.method_by_name("get_AbsoluteUri").is_some());
        assert!(iface.method_by_name("get_DisplayUri").is_some());
        assert!(iface.method_by_name("nonexistent").is_none());
    }

    // -----------------------------------------------------------------------
    // IID / signature computation
    // -----------------------------------------------------------------------

    #[test]
    fn iid_interface() {
        let table = MetadataTable::new();
        let iid = GUID::from_u128(0x12345678_1234_1234_1234_123456789abc);
        assert_eq!(table.interface(iid).iid(), Some(iid));
    }

    #[test]
    fn iid_parameterized_async_operation() {
        let table = MetadataTable::new();
        let g = table.generic(IASYNC_OPERATION, 1);
        let p = table.parameterized(&g, &[table.hstring()]);

        // Must match the IID computed by windows_future for IAsyncOperation<HSTRING>
        assert_eq!(
            p.iid().unwrap(),
            windows_future::IAsyncOperation::<windows_core::HSTRING>::IID,
        );
    }

    #[test]
    fn iid_runtime_class_as_type_arg() {
        let table = MetadataTable::new();
        let storage_file = table.runtime_class(
            "Windows.Storage.StorageFile".into(),
            GUID::from_u128(0xFA3F6186_4214_428C_A64C_14C9AC7315EA),
        );
        let g = table.generic(IASYNC_OPERATION, 1);
        let ty = table.parameterized(&g, &[storage_file]);

        let expected_iid = GUID::from_u128(0x5e52f8ce_aced_5a42_95b4_f674dd84885e);
        assert_eq!(ty.iid().unwrap(), expected_iid);
    }

    #[test]
    fn signature_string() {
        let table = MetadataTable::new();
        assert_eq!(table.i32_type().signature_string(), "i4");
        assert_eq!(table.hstring().signature_string(), "string");

        let g = table.generic(IASYNC_OPERATION, 1);
        let sig = table.parameterized(&g, &[table.hstring()]);
        assert_eq!(
            sig.signature_string(),
            "pinterface({9fc2b0bb-e446-44e2-aa61-9cab8f636af2};string)",
        );
    }

    #[test]
    fn guid_braced_format() {
        let guid = GUID::from_u128(0x9fc2b0bb_e446_44e2_aa61_9cab8f636af2);
        assert_eq!(format_guid_braced(&guid), "{9fc2b0bb-e446-44e2-aa61-9cab8f636af2}");
    }

    // -----------------------------------------------------------------------
    // End-to-end: register → invoke → verify (requires WinRT runtime)
    // -----------------------------------------------------------------------

    #[test]
    fn e2e_uri_create_and_query() {
        use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};
        use windows_core::{Interface, h};

        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };
        let table = MetadataTable::new();

        // Register interfaces
        let factory_iid = GUID::from_u128(0x44A9796F_723E_4FDF_A218_033E75B0C084);
        let factory_iface = table.register_interface("IUriRuntimeClassFactory", factory_iid)
            .add_method("CreateUri", MethodSignature::new(&table)
                .add_in(table.hstring()).add_out(table.object()));

        let uri_iid = GUID::from_u128(0x9E365E57_48B2_4160_956F_C7385120BBFC);
        let uri_iface = table.register_interface("IUriRuntimeClass", uri_iid)
            .add_method("get_AbsoluteUri", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_DisplayUri", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_Domain", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_Extension", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_Fragment", MethodSignature::new(&table).add_out(table.hstring()))
            .add_method("get_Host", MethodSignature::new(&table).add_out(table.hstring()));

        // Activate factory and QI
        let factory = unsafe {
            windows::Win32::System::WinRT::RoGetActivationFactory::<
                windows::Win32::System::WinRT::IActivationFactory,
            >(h!("Windows.Foundation.Uri"))
        }.unwrap();
        let mut factory_ptr = std::ptr::null_mut();
        unsafe { factory.cast::<windows_core::IUnknown>().unwrap()
            .query(&factory_iid, &mut factory_ptr).ok().unwrap(); }

        // CreateUri via method_by_name
        let uri_val = factory_iface.method_by_name("CreateUri").unwrap()
            .invoke(factory_ptr, &[
                WinRTValue::HString(windows_core::HSTRING::from("https://www.example.com/path?q=1#frag"))
            ]).unwrap();
        let uri_obj = uri_val[0].as_object().unwrap();
        let mut uri_ptr = std::ptr::null_mut();
        unsafe { uri_obj.query(&uri_iid, &mut uri_ptr).ok().unwrap(); }

        // get_Host via method_by_name
        let host = uri_iface.method_by_name("get_Host").unwrap()
            .invoke(uri_ptr, &[]).unwrap()[0].as_hstring().unwrap();
        assert_eq!(host.to_string(), "www.example.com");

        // get_AbsoluteUri via vtable index
        let abs_uri = uri_iface.method(6).unwrap()
            .invoke(uri_ptr, &[]).unwrap()[0].as_hstring().unwrap();
        assert_eq!(abs_uri.to_string(), "https://www.example.com/path?q=1#frag");

        // get_Domain via method_by_name
        let domain = uri_iface.method_by_name("get_Domain").unwrap()
            .invoke(uri_ptr, &[]).unwrap()[0].as_hstring().unwrap();
        assert_eq!(domain.to_string(), "example.com");
    }

    #[test]
    fn e2e_geopoint_struct_in_param() -> windows::core::Result<()> {
        use windows::Devices::Geolocation::{Geopoint, IGeopointFactory};
        use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};
        use windows_core::{Interface, h};

        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };
        let table = MetadataTable::new();

        // Register BasicGeoposition struct
        let f64_h = table.f64_type();
        let geo_type = table.struct_type(
            "Windows.Devices.Geolocation.BasicGeoposition",
            &[f64_h.clone(), f64_h.clone(), f64_h],
        );

        // Register IGeopointFactory
        let factory_iid = IGeopointFactory::IID;
        let factory_iface = table.register_interface("IGeopointFactory", factory_iid)
            .add_method("Create", MethodSignature::new(&table)
                .add_in(geo_type.clone()).add_out(table.object()));

        // Create struct value
        let mut geo_val = geo_type.default_value();
        geo_val.set_field(0, 47.643f64);   // Latitude
        geo_val.set_field(1, -122.131f64);  // Longitude
        geo_val.set_field(2, 100.0f64);     // Altitude

        // Activate and call
        let af = unsafe {
            windows::Win32::System::WinRT::RoGetActivationFactory::<
                windows::Win32::System::WinRT::IActivationFactory,
            >(h!("Windows.Devices.Geolocation.Geopoint"))
        }?;
        let mut factory_ptr = std::ptr::null_mut();
        unsafe { af.cast::<windows_core::IUnknown>().unwrap()
            .query(&factory_iid, &mut factory_ptr).ok().unwrap(); }

        let result = factory_iface.method_by_name("Create").unwrap()
            .invoke(factory_ptr, &[WinRTValue::Struct(geo_val)])
            .map_err(|e| match e { crate::result::Error::WindowsError(we) => we, _ => panic!("{:?}", e) })?;

        // Verify via static projection
        let geopoint: Geopoint = result[0].as_object().unwrap().cast()?;
        let pos = geopoint.Position()?;
        assert!((pos.Latitude - 47.643).abs() < 1e-6);
        assert!((pos.Longitude - (-122.131)).abs() < 1e-6);
        assert!((pos.Altitude - 100.0).abs() < 1e-6);
        Ok(())
    }
}
