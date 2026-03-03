use std::alloc::Layout;
use std::collections::HashMap;
use std::sync::{Arc, RwLock};

use windows_core::GUID;

use crate::signature::{Method, MethodSignature};
use crate::types::WinRTType;
use crate::value::WinRTValue;

// ===========================================================================
// Struct layout engine (formerly registry.rs)
// ===========================================================================

/// Primitive types that can appear as fields in WinRT value types.
///
/// ## Not yet supported field types
///
/// - **Enum**: WinRT enums are either I32 (signed) or U32 (unsigned) at the ABI level.
///   Can be represented using existing I32/U32 primitives, but enum-specific semantics
///   (named values, type identity) are not preserved.
///
/// - **HSTRING**: A reference-counted string handle (8 bytes, pointer). Requires special
///   Drop (WindowsDeleteString), Clone (WindowsDuplicateString), and set_field (release old)
///   semantics. Only ~5 WinRT structs use String fields (e.g. AccessListEntry, XmlnsDefinition).
///
/// - **IReference\<T\>**: A nullable value type represented as a COM interface pointer (8 bytes).
///   Same reference-counting challenges as HSTRING. Only appears in HttpProgress
///   (`TotalBytesToSend: IReference<u64>`).
///
/// HSTRING and IReference would make the struct "non-blittable" — ValueTypeData's
/// Drop/Clone/set_field currently assume pure memcpy semantics and would need per-field
/// lifecycle management.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PrimitiveType {
    Bool,
    I8,
    U8,
    I16,
    U16,
    Char16,
    I32,
    U32,
    I64,
    U64,
    F32,
    F64,
    Guid,
}

impl PrimitiveType {
    pub fn size_of(self) -> usize {
        match self {
            PrimitiveType::Bool | PrimitiveType::I8 | PrimitiveType::U8 => 1,
            PrimitiveType::I16 | PrimitiveType::U16 | PrimitiveType::Char16 => 2,
            PrimitiveType::I32 | PrimitiveType::U32 | PrimitiveType::F32 => 4,
            PrimitiveType::I64 | PrimitiveType::U64 | PrimitiveType::F64 => 8,
            PrimitiveType::Guid => 16,
        }
    }

    pub fn align_of(self) -> usize {
        match self {
            // Guid is { u32, u16, u16, [u8; 8] } — alignment follows the largest field (u32)
            PrimitiveType::Guid => 4,
            _ => self.size_of(),
        }
    }

    fn libffi_type(self) -> libffi::middle::Type {
        use libffi::middle::Type;
        match self {
            PrimitiveType::Bool => Type::u8(),
            PrimitiveType::I8 => Type::i8(),
            PrimitiveType::U8 => Type::u8(),
            PrimitiveType::I16 => Type::i16(),
            PrimitiveType::U16 | PrimitiveType::Char16 => Type::u16(),
            PrimitiveType::I32 => Type::i32(),
            PrimitiveType::U32 => Type::u32(),
            PrimitiveType::I64 => Type::i64(),
            PrimitiveType::U64 => Type::u64(),
            PrimitiveType::F32 => Type::f32(),
            PrimitiveType::F64 => Type::f64(),
            PrimitiveType::Guid => Type::structure(vec![
                Type::u32(), Type::u16(), Type::u16(),
                Type::u8(), Type::u8(), Type::u8(), Type::u8(),
                Type::u8(), Type::u8(), Type::u8(), Type::u8(),
            ]),
        }
    }
}

/// Internal type identifier. Not exposed publicly — users only see `TypeHandle`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum TypeKind {
    Primitive(PrimitiveType),
    Struct(u32),
}

/// Internal struct data stored in the registry.
struct StructEntry {
    field_kinds: Vec<TypeKind>,
    field_offsets: Vec<usize>,
    layout: Layout,
}

/// Registry of value types. Always lives behind `Arc`, supports concurrent reads
/// and append-only mutation via `RwLock`.
pub struct TypeRegistry {
    structs: RwLock<Vec<StructEntry>>,
}

impl TypeRegistry {
    pub fn new() -> Arc<Self> {
        Arc::new(TypeRegistry {
            structs: RwLock::new(Vec::new()),
        })
    }

    pub fn primitive(self: &Arc<Self>, ty: PrimitiveType) -> TypeHandle {
        TypeHandle {
            registry: Arc::clone(self),
            kind: TypeKind::Primitive(ty),
        }
    }

    pub fn define_struct(self: &Arc<Self>, fields: &[TypeHandle]) -> TypeHandle {
        let field_kinds: Vec<TypeKind> = fields.iter().map(|h| h.kind).collect();
        let (field_offsets, layout) = self.compute_layout(&field_kinds);
        let mut structs = self.structs.write().unwrap();
        let id = structs.len() as u32;
        structs.push(StructEntry {
            field_kinds,
            field_offsets,
            layout,
        });
        TypeHandle {
            registry: Arc::clone(self),
            kind: TypeKind::Struct(id),
        }
    }

    fn size_of_kind(&self, kind: TypeKind) -> usize {
        match kind {
            TypeKind::Primitive(p) => p.size_of(),
            TypeKind::Struct(id) => self.structs.read().unwrap()[id as usize].layout.size(),
        }
    }

    fn align_of_kind(&self, kind: TypeKind) -> usize {
        match kind {
            TypeKind::Primitive(p) => p.align_of(),
            TypeKind::Struct(id) => self.structs.read().unwrap()[id as usize].layout.align(),
        }
    }

    fn layout_of_kind(&self, kind: TypeKind) -> Layout {
        match kind {
            TypeKind::Primitive(p) => Layout::from_size_align(p.size_of(), p.align_of()).unwrap(),
            TypeKind::Struct(id) => self.structs.read().unwrap()[id as usize].layout,
        }
    }

    fn field_count_kind(&self, kind: TypeKind) -> usize {
        match kind {
            TypeKind::Primitive(_) => panic!("Primitive types have no fields"),
            TypeKind::Struct(id) => self.structs.read().unwrap()[id as usize].field_kinds.len(),
        }
    }

    fn field_offset_kind(&self, kind: TypeKind, index: usize) -> usize {
        match kind {
            TypeKind::Primitive(_) => panic!("Primitive types have no fields"),
            TypeKind::Struct(id) => self.structs.read().unwrap()[id as usize].field_offsets[index],
        }
    }

    fn field_kind(&self, kind: TypeKind, index: usize) -> TypeKind {
        match kind {
            TypeKind::Primitive(_) => panic!("Primitive types have no fields"),
            TypeKind::Struct(id) => self.structs.read().unwrap()[id as usize].field_kinds[index],
        }
    }

    fn libffi_type_kind(&self, kind: TypeKind) -> libffi::middle::Type {
        match kind {
            TypeKind::Primitive(p) => p.libffi_type(),
            TypeKind::Struct(id) => {
                let structs = self.structs.read().unwrap();
                let field_types: Vec<libffi::middle::Type> = structs[id as usize]
                    .field_kinds
                    .iter()
                    .map(|f| self.libffi_type_kind(*f))
                    .collect();
                libffi::middle::Type::structure(field_types)
            }
        }
    }

    fn compute_layout(&self, fields: &[TypeKind]) -> (Vec<usize>, Layout) {
        let mut offsets = Vec::with_capacity(fields.len());
        let mut layout = Layout::from_size_align(0, 1).unwrap();

        for field in fields {
            let field_layout = self.layout_of_kind(*field);
            let (new_layout, offset) = layout.extend(field_layout).unwrap();
            offsets.push(offset);
            layout = new_layout;
        }

        (offsets, layout.pad_to_align())
    }
}

/// A handle to a type in the registry. Carries an `Arc<TypeRegistry>` so it
/// can query layout and create values without needing a separate registry reference.
#[derive(Clone)]
pub struct TypeHandle {
    registry: Arc<TypeRegistry>,
    kind: TypeKind,
}

impl std::fmt::Debug for TypeHandle {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("TypeHandle")
            .field("kind", &self.kind)
            .field("size", &self.size_of())
            .field("align", &self.align_of())
            .finish()
    }
}

impl PartialEq for TypeHandle {
    fn eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.registry, &other.registry) && self.kind == other.kind
    }
}

impl Eq for TypeHandle {}

impl TypeHandle {
    pub fn size_of(&self) -> usize {
        self.registry.size_of_kind(self.kind)
    }

    pub fn align_of(&self) -> usize {
        self.registry.align_of_kind(self.kind)
    }

    pub fn layout(&self) -> Layout {
        self.registry.layout_of_kind(self.kind)
    }

    pub fn libffi_type(&self) -> libffi::middle::Type {
        self.registry.libffi_type_kind(self.kind)
    }

    pub fn field_count(&self) -> usize {
        self.registry.field_count_kind(self.kind)
    }

    pub fn field_offset(&self, index: usize) -> usize {
        self.registry.field_offset_kind(self.kind, index)
    }

    pub fn field_type(&self, index: usize) -> TypeHandle {
        TypeHandle {
            registry: Arc::clone(&self.registry),
            kind: self.registry.field_kind(self.kind, index),
        }
    }

    pub fn default_value(&self) -> ValueTypeData {
        ValueTypeData::new(self)
    }
}

/// A dynamically-typed value matching a struct layout from the registry.
///
/// Owns an aligned heap allocation. Holds a `TypeHandle` internally so
/// field access methods are self-contained.
pub struct ValueTypeData {
    type_handle: TypeHandle,
    ptr: *mut u8,
}

impl std::fmt::Debug for ValueTypeData {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("ValueTypeData")
            .field("type_handle", &self.type_handle)
            .field("ptr", &self.ptr)
            .finish()
    }
}

impl ValueTypeData {
    fn new(handle: &TypeHandle) -> Self {
        let layout = handle.layout();
        let ptr = if layout.size() > 0 {
            unsafe { std::alloc::alloc_zeroed(layout) }
        } else {
            std::ptr::null_mut()
        };
        Self {
            type_handle: handle.clone(),
            ptr,
        }
    }

    pub fn type_handle(&self) -> &TypeHandle {
        &self.type_handle
    }

    pub fn as_ptr(&self) -> *const u8 {
        self.ptr
    }

    pub fn as_mut_ptr(&mut self) -> *mut u8 {
        self.ptr
    }

    pub fn get_field<T: Copy>(&self, index: usize) -> T {
        let h = &self.type_handle;
        let offset = h.field_offset(index);
        assert_eq!(
            std::mem::size_of::<T>(),
            h.field_type(index).size_of(),
            "get_field<T> size mismatch"
        );
        unsafe { (self.ptr.add(offset) as *const T).read() }
    }

    pub fn set_field<T: Copy>(&mut self, index: usize, value: T) {
        let h = &self.type_handle;
        let offset = h.field_offset(index);
        assert_eq!(
            std::mem::size_of::<T>(),
            h.field_type(index).size_of(),
            "set_field<T> size mismatch"
        );
        unsafe { (self.ptr.add(offset) as *mut T).write(value) }
    }

    pub fn get_field_struct(&self, index: usize) -> ValueTypeData {
        let h = &self.type_handle;
        let offset = h.field_offset(index);
        let field_handle = h.field_type(index);
        let layout = field_handle.layout();
        let result = field_handle.default_value();
        if layout.size() > 0 {
            unsafe {
                std::ptr::copy_nonoverlapping(
                    self.ptr.add(offset),
                    result.ptr,
                    layout.size(),
                );
            }
        }
        result
    }

    pub fn set_field_struct(&mut self, index: usize, value: &ValueTypeData) {
        let h = &self.type_handle;
        let offset = h.field_offset(index);
        let field_handle = h.field_type(index);
        let size = field_handle.size_of();
        assert_eq!(
            size,
            value.type_handle.size_of(),
            "set_field_struct size mismatch"
        );
        if size > 0 {
            unsafe {
                std::ptr::copy_nonoverlapping(
                    value.ptr,
                    self.ptr.add(offset),
                    size,
                );
            }
        }
    }

    pub fn call_method_struct_to_object(
        &self,
        obj_raw: *mut std::ffi::c_void,
        method_index: usize,
    ) -> windows_core::Result<windows_core::IUnknown> {
        use crate::call::get_vtable_function_ptr;
        use libffi::middle::{arg, Cif, CodePtr, Type};
        use windows_core::Interface;

        let fptr = get_vtable_function_ptr(obj_raw, method_index);
        let cif = Cif::new(
            vec![
                Type::pointer(),
                self.type_handle.libffi_type(),
                Type::pointer(),
            ],
            Type::i32(),
        );

        let mut out: *mut std::ffi::c_void = std::ptr::null_mut();
        let data_ref = unsafe { &*self.ptr };
        let hr: windows_core::HRESULT = unsafe {
            cif.call(
                CodePtr(fptr),
                &[arg(&obj_raw), arg(data_ref), arg(&(&mut out))],
            )
        };
        hr.ok()?;
        Ok(unsafe { windows_core::IUnknown::from_raw(out as _) })
    }
}

impl Drop for ValueTypeData {
    fn drop(&mut self) {
        let layout = self.type_handle.layout();
        if layout.size() > 0 {
            unsafe { std::alloc::dealloc(self.ptr, layout) }
        }
    }
}

impl Clone for ValueTypeData {
    fn clone(&self) -> Self {
        let layout = self.type_handle.layout();
        if layout.size() == 0 {
            return Self {
                type_handle: self.type_handle.clone(),
                ptr: std::ptr::null_mut(),
            };
        }
        let ptr = unsafe {
            let p = std::alloc::alloc(layout);
            std::ptr::copy_nonoverlapping(self.ptr, p, layout.size());
            p
        };
        Self {
            type_handle: self.type_handle.clone(),
            ptr,
        }
    }
}

// ---------------------------------------------------------------------------
// Entry types
// ---------------------------------------------------------------------------

/// A single WinRT type registered in the TypeTable.
#[derive(Debug)]
pub enum TypeEntry {
    Interface(InterfaceEntry),
    RuntimeClass(RuntimeClassEntry),
    Struct(StructTypeEntry),
    Enum(EnumEntry),
    Delegate(DelegateEntry),
}

/// A WinRT interface with pre-built Method objects (cached Cifs).
#[derive(Debug)]
pub struct InterfaceEntry {
    pub name: String,
    pub iid: GUID,
    pub methods: Vec<MethodEntry>,
}

/// A single method on an interface, with its pre-built Cif.
#[derive(Debug)]
pub struct MethodEntry {
    pub name: String,
    pub method: Method,
}

/// A WinRT runtime class.
#[derive(Debug)]
pub struct RuntimeClassEntry {
    pub name: String,
    pub default_iid: GUID,
    pub interfaces: Vec<InterfaceRef>,
    pub has_default_constructor: bool,
}

/// A reference to an interface from a runtime class.
#[derive(Debug, Clone)]
pub struct InterfaceRef {
    pub name: String,
    pub role: InterfaceRole,
}

/// The role of an interface relative to a runtime class.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InterfaceRole {
    Default,
    Factory,
    Static,
    Other,
}

/// A WinRT struct (value type), backed by the TypeRegistry.
#[derive(Debug)]
pub struct StructTypeEntry {
    pub name: String,
    pub type_handle: TypeHandle,
    pub field_names: Vec<String>,
    pub field_types: Vec<WinRTType>,
}

/// A WinRT enum with named members.
#[derive(Debug)]
pub struct EnumEntry {
    pub name: String,
    pub underlying: WinRTType,
    pub members: Vec<(String, i32)>,
}

/// A WinRT delegate type.
#[derive(Debug)]
pub struct DelegateEntry {
    pub name: String,
    pub iid: GUID,
}

// ---------------------------------------------------------------------------
// TypeTable
// ---------------------------------------------------------------------------

/// Centralized registry of WinRT types, indexed by full name.
///
/// Thread-safe, append-only, Arc-wrapped for shared ownership.
/// Composes `TypeRegistry` for struct layout computation and caches
/// pre-built `Cif` objects for interface method calls.
pub struct TypeTable {
    entries: RwLock<HashMap<String, TypeEntry>>,
    struct_registry: Arc<TypeRegistry>,
}

impl TypeTable {
    /// Create a new empty TypeTable with its own TypeRegistry.
    pub fn new() -> Arc<Self> {
        Arc::new(TypeTable {
            entries: RwLock::new(HashMap::new()),
            struct_registry: TypeRegistry::new(),
        })
    }

    /// Create a TypeTable sharing an existing TypeRegistry.
    pub fn with_registry(registry: Arc<TypeRegistry>) -> Arc<Self> {
        Arc::new(TypeTable {
            entries: RwLock::new(HashMap::new()),
            struct_registry: registry,
        })
    }

    /// Access the underlying TypeRegistry.
    pub fn struct_registry(&self) -> &Arc<TypeRegistry> {
        &self.struct_registry
    }

    // -----------------------------------------------------------------------
    // Registration API
    // -----------------------------------------------------------------------

    /// Register an interface type with pre-built methods.
    ///
    /// `methods` is a list of `(method_name, MethodSignature)` pairs.
    /// Cifs are built during registration, starting at vtable index
    /// `base_vtable_index` (typically 6 for IInspectable-derived interfaces).
    pub fn register_interface(
        &self,
        name: &str,
        iid: GUID,
        base_vtable_index: usize,
        methods: Vec<(String, MethodSignature)>,
    ) {
        let built_methods: Vec<MethodEntry> = methods
            .into_iter()
            .enumerate()
            .map(|(i, (method_name, sig))| MethodEntry {
                name: method_name,
                method: sig.build(base_vtable_index + i),
            })
            .collect();

        let entry = TypeEntry::Interface(InterfaceEntry {
            name: name.to_string(),
            iid,
            methods: built_methods,
        });
        self.entries.write().unwrap().insert(name.to_string(), entry);
    }

    /// Register a runtime class.
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

    /// Register a struct type. Returns the TypeHandle for immediate use.
    ///
    /// If a struct with the same name is already registered, returns
    /// the existing TypeHandle (deduplication).
    pub fn register_struct(
        &self,
        name: &str,
        fields: Vec<(String, WinRTType)>,
    ) -> TypeHandle {
        // Check for existing registration (dedup)
        {
            let entries = self.entries.read().unwrap();
            if let Some(TypeEntry::Struct(s)) = entries.get(name) {
                return s.type_handle.clone();
            }
        }

        let (field_names, field_types): (Vec<String>, Vec<WinRTType>) =
            fields.into_iter().unzip();

        let field_handles: Vec<TypeHandle> = field_types
            .iter()
            .map(|t| self.winrt_type_to_handle(t))
            .collect();

        let type_handle = self.struct_registry.define_struct(&field_handles);

        let entry = TypeEntry::Struct(StructTypeEntry {
            name: name.to_string(),
            type_handle: type_handle.clone(),
            field_names,
            field_types,
        });
        self.entries.write().unwrap().insert(name.to_string(), entry);
        type_handle
    }

    /// Register an enum type.
    pub fn register_enum(
        &self,
        name: &str,
        underlying: WinRTType,
        members: Vec<(String, i32)>,
    ) {
        let entry = TypeEntry::Enum(EnumEntry {
            name: name.to_string(),
            underlying,
            members,
        });
        self.entries.write().unwrap().insert(name.to_string(), entry);
    }

    /// Register a delegate type.
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

    /// Check if a type is registered.
    pub fn contains(&self, name: &str) -> bool {
        self.entries.read().unwrap().contains_key(name)
    }

    /// Get the IID for a type by name.
    ///
    /// Returns Some for Interface, Delegate, and RuntimeClass (default IID).
    /// Returns None for Struct, Enum, or unknown types.
    pub fn get_iid(&self, name: &str) -> Option<GUID> {
        let entries = self.entries.read().unwrap();
        match entries.get(name)? {
            TypeEntry::Interface(e) => Some(e.iid),
            TypeEntry::Delegate(e) => Some(e.iid),
            TypeEntry::RuntimeClass(e) => Some(e.default_iid),
            _ => None,
        }
    }

    /// Get the struct TypeHandle by name.
    pub fn get_struct_handle(&self, name: &str) -> Option<TypeHandle> {
        let entries = self.entries.read().unwrap();
        match entries.get(name)? {
            TypeEntry::Struct(e) => Some(e.type_handle.clone()),
            _ => None,
        }
    }

    /// Get the WinRTType descriptor for a registered type.
    pub fn get_winrt_type(&self, name: &str) -> Option<WinRTType> {
        let entries = self.entries.read().unwrap();
        match entries.get(name)? {
            TypeEntry::Interface(e) => Some(WinRTType::Interface(e.iid)),
            TypeEntry::RuntimeClass(e) => {
                Some(WinRTType::RuntimeClass(e.name.clone(), e.default_iid))
            }
            TypeEntry::Struct(e) => Some(WinRTType::Struct(e.type_handle.clone())),
            TypeEntry::Enum(e) => Some(e.underlying.clone()),
            TypeEntry::Delegate(e) => Some(WinRTType::Delegate(e.iid)),
        }
    }

    /// Get an enum member value by enum name and member name.
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

    /// Call a method on a COM object by interface name and method name.
    ///
    /// Uses the pre-built Cif from registration time — no Cif rebuild overhead.
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
        let method_entry = iface
            .methods
            .iter()
            .find(|m| m.name == method_name)
            .ok_or_else(|| {
                crate::result::Error::MethodNotFound(
                    interface_name.to_string(),
                    method_name.to_string(),
                )
            })?;
        method_entry
            .method
            .call_dynamic(obj, args)
            .map_err(|e| crate::result::Error::WindowsError(e))
    }

    /// Call a method by interface name and method index (0-based,
    /// relative to the first WinRT method registered for this interface).
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
        let method_entry = iface.methods.get(method_index).ok_or_else(|| {
            crate::result::Error::MethodNotFound(
                interface_name.to_string(),
                format!("[index {}]", method_index),
            )
        })?;
        method_entry
            .method
            .call_dynamic(obj, args)
            .map_err(|e| crate::result::Error::WindowsError(e))
    }

    // -----------------------------------------------------------------------
    // Internal helpers
    // -----------------------------------------------------------------------

    /// Convert a WinRTType (used for struct fields) to a TypeHandle.
    fn winrt_type_to_handle(&self, typ: &WinRTType) -> TypeHandle {
        match typ {
            WinRTType::Bool => self.struct_registry.primitive(PrimitiveType::Bool),
            WinRTType::I8 => self.struct_registry.primitive(PrimitiveType::I8),
            WinRTType::U8 => self.struct_registry.primitive(PrimitiveType::U8),
            WinRTType::I16 => self.struct_registry.primitive(PrimitiveType::I16),
            WinRTType::U16 | WinRTType::Char16 => {
                self.struct_registry.primitive(PrimitiveType::U16)
            }
            WinRTType::I32 => self.struct_registry.primitive(PrimitiveType::I32),
            WinRTType::U32 => self.struct_registry.primitive(PrimitiveType::U32),
            WinRTType::I64 => self.struct_registry.primitive(PrimitiveType::I64),
            WinRTType::U64 => self.struct_registry.primitive(PrimitiveType::U64),
            WinRTType::F32 => self.struct_registry.primitive(PrimitiveType::F32),
            WinRTType::F64 => self.struct_registry.primitive(PrimitiveType::F64),
            WinRTType::Guid => self.struct_registry.primitive(PrimitiveType::Guid),
            WinRTType::Struct(handle) => handle.clone(),
            _ => panic!("Unsupported struct field type: {:?}", typ),
        }
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    // -----------------------------------------------------------------------
    // Struct layout tests (formerly in registry.rs)
    // -----------------------------------------------------------------------

    #[test]
    fn primitive_sizes() {
        let reg = TypeRegistry::new();
        let f32_h = reg.primitive(PrimitiveType::F32);
        let f64_h = reg.primitive(PrimitiveType::F64);
        let i32_h = reg.primitive(PrimitiveType::I32);
        let u8_h = reg.primitive(PrimitiveType::U8);

        assert_eq!(f32_h.size_of(), 4);
        assert_eq!(f64_h.size_of(), 8);
        assert_eq!(i32_h.size_of(), 4);
        assert_eq!(u8_h.size_of(), 1);

        assert_eq!(f32_h.align_of(), 4);
        assert_eq!(f64_h.align_of(), 8);
    }

    #[test]
    fn point_layout() {
        let reg = TypeRegistry::new();
        let f32_h = reg.primitive(PrimitiveType::F32);
        let point = reg.define_struct(&[f32_h.clone(), f32_h]);

        assert_eq!(point.size_of(), 8);
        assert_eq!(point.align_of(), 4);
        assert_eq!(point.field_count(), 2);
        assert_eq!(point.field_offset(0), 0);
        assert_eq!(point.field_offset(1), 4);
    }

    #[test]
    fn rect_layout() {
        let reg = TypeRegistry::new();
        let f32_h = reg.primitive(PrimitiveType::F32);
        let rect = reg.define_struct(&[f32_h.clone(), f32_h.clone(), f32_h.clone(), f32_h]);

        assert_eq!(rect.size_of(), 16);
        assert_eq!(rect.align_of(), 4);
        assert_eq!(rect.field_offset(0), 0);
        assert_eq!(rect.field_offset(1), 4);
        assert_eq!(rect.field_offset(2), 8);
        assert_eq!(rect.field_offset(3), 12);
    }

    #[test]
    fn basic_geoposition_layout() {
        let reg = TypeRegistry::new();
        let f64_h = reg.primitive(PrimitiveType::F64);
        let geo = reg.define_struct(&[f64_h.clone(), f64_h.clone(), f64_h]);

        assert_eq!(geo.size_of(), 24);
        assert_eq!(geo.align_of(), 8);
        assert_eq!(geo.field_offset(0), 0);
        assert_eq!(geo.field_offset(1), 8);
        assert_eq!(geo.field_offset(2), 16);
    }

    #[test]
    fn mixed_field_alignment() {
        let reg = TypeRegistry::new();
        let u8_h = reg.primitive(PrimitiveType::U8);
        let i32_h = reg.primitive(PrimitiveType::I32);
        let s = reg.define_struct(&[u8_h.clone(), i32_h, u8_h]);

        assert_eq!(s.size_of(), 12);
        assert_eq!(s.align_of(), 4);
        assert_eq!(s.field_offset(0), 0);
        assert_eq!(s.field_offset(1), 4);
        assert_eq!(s.field_offset(2), 8);
    }

    #[test]
    fn nested_struct_layout() {
        let reg = TypeRegistry::new();
        let f32_h = reg.primitive(PrimitiveType::F32);
        let f64_h = reg.primitive(PrimitiveType::F64);
        let inner = reg.define_struct(&[f32_h.clone(), f32_h]);
        let outer = reg.define_struct(&[inner, f64_h]);

        assert_eq!(outer.size_of(), 16);
        assert_eq!(outer.align_of(), 8);
        assert_eq!(outer.field_offset(0), 0);
        assert_eq!(outer.field_offset(1), 8);
    }

    #[test]
    fn value_type_data_field_access() {
        let reg = TypeRegistry::new();
        let f32_h = reg.primitive(PrimitiveType::F32);
        let point = reg.define_struct(&[f32_h.clone(), f32_h]);

        let mut val = point.default_value();
        val.set_field(0, 10.0f32);
        val.set_field(1, 20.0f32);

        assert_eq!(val.get_field::<f32>(0), 10.0);
        assert_eq!(val.get_field::<f32>(1), 20.0);
    }

    #[test]
    fn nested_struct_field_access() {
        let reg = TypeRegistry::new();
        let f32_h = reg.primitive(PrimitiveType::F32);
        let point_type = reg.define_struct(&[f32_h.clone(), f32_h.clone()]);
        let ink_type = reg.define_struct(&[point_type.clone(), f32_h]);

        let mut point_val = point_type.default_value();
        point_val.set_field(0, 10.0f32);
        point_val.set_field(1, 20.0f32);

        let mut ink_val = ink_type.default_value();
        ink_val.set_field_struct(0, &point_val);
        ink_val.set_field(1, 5.0f32);

        let read_point = ink_val.get_field_struct(0);
        assert_eq!(read_point.get_field::<f32>(0), 10.0);
        assert_eq!(read_point.get_field::<f32>(1), 20.0);
        assert_eq!(ink_val.get_field::<f32>(1), 5.0);
    }

    #[test]
    fn value_type_data_clone() {
        let reg = TypeRegistry::new();
        let f64_h = reg.primitive(PrimitiveType::F64);
        let geo = reg.define_struct(&[f64_h.clone(), f64_h.clone(), f64_h]);

        let mut val = geo.default_value();
        val.set_field(0, 47.6f64);
        val.set_field(1, -122.3f64);
        val.set_field(2, 100.0f64);

        let cloned = val.clone();
        assert_eq!(cloned.get_field::<f64>(0), 47.6);
        assert_eq!(cloned.get_field::<f64>(1), -122.3);
        assert_eq!(cloned.get_field::<f64>(2), 100.0);

        val.set_field(0, 0.0f64);
        assert_eq!(cloned.get_field::<f64>(0), 47.6);
    }

    #[test]
    fn value_type_matches_windows_point_layout() {
        use windows::Foundation::Point;

        let reg = TypeRegistry::new();
        let f32_h = reg.primitive(PrimitiveType::F32);
        let point = reg.define_struct(&[f32_h.clone(), f32_h]);

        assert_eq!(point.size_of(), std::mem::size_of::<Point>());
        assert_eq!(point.align_of(), std::mem::align_of::<Point>());

        let mut val = point.default_value();
        val.set_field(0, 10.0f32);
        val.set_field(1, 20.0f32);

        let win_point: &Point = unsafe { &*(val.as_ptr() as *const Point) };
        assert_eq!(win_point.X, 10.0);
        assert_eq!(win_point.Y, 20.0);
    }

    #[test]
    fn libffi_type_primitive() {
        let reg = TypeRegistry::new();
        let _ = reg.primitive(PrimitiveType::F32).libffi_type();
        let _ = reg.primitive(PrimitiveType::I64).libffi_type();
    }

    #[test]
    fn libffi_type_struct() {
        let reg = TypeRegistry::new();
        let f32_h = reg.primitive(PrimitiveType::F32);
        let f64_h = reg.primitive(PrimitiveType::F64);
        let point = reg.define_struct(&[f32_h.clone(), f32_h]);
        let _ = point.libffi_type();

        let outer = reg.define_struct(&[point, f64_h]);
        let _ = outer.libffi_type();
    }

    #[test]
    fn geopoint_create_via_registry() -> windows::core::Result<()> {
        use libffi::middle::{Cif, CodePtr, arg};
        use windows::Devices::Geolocation::Geopoint;
        use windows::Win32::System::WinRT::{
            IActivationFactory, RO_INIT_MULTITHREADED, RoGetActivationFactory, RoInitialize,
        };
        use windows::core::{Interface, h};
        use windows_core::HRESULT;

        use crate::call::get_vtable_function_ptr;

        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };

        let reg = TypeRegistry::new();
        let f64_h = reg.primitive(PrimitiveType::F64);
        let geo_type = reg.define_struct(&[f64_h.clone(), f64_h.clone(), f64_h]);

        let mut geo_val = geo_type.default_value();
        geo_val.set_field(0, 47.643f64);
        geo_val.set_field(1, -122.131f64);
        geo_val.set_field(2, 100.0f64);

        let afactory = unsafe {
            RoGetActivationFactory::<IActivationFactory>(
                h!("Windows.Devices.Geolocation.Geopoint"),
            )
        }?;
        let geopoint_factory =
            afactory.cast::<windows::Devices::Geolocation::IGeopointFactory>()?;
        let fptr = get_vtable_function_ptr(geopoint_factory.as_raw(), 6);

        let cif = Cif::new(
            vec![
                libffi::middle::Type::pointer(),
                geo_type.libffi_type(),
                libffi::middle::Type::pointer(),
            ],
            libffi::middle::Type::i32(),
        );

        let mut out: *mut std::ffi::c_void = std::ptr::null_mut();
        let this = geopoint_factory.as_raw();
        let geo_data_ref = unsafe { &*geo_val.as_ptr() };
        let hr: HRESULT = unsafe {
            cif.call(
                CodePtr(fptr),
                &[arg(&this), arg(geo_data_ref), arg(&(&mut out))],
            )
        };
        hr.ok()?;

        let geopoint = unsafe { Geopoint::from_raw(out) };
        let pos = geopoint.Position()?;
        assert!((pos.Latitude - 47.643).abs() < 1e-6);
        assert!((pos.Longitude - (-122.131)).abs() < 1e-6);
        assert!((pos.Altitude - 100.0).abs() < 1e-6);

        Ok(())
    }

    // -----------------------------------------------------------------------
    // TypeTable tests
    // -----------------------------------------------------------------------

    #[test]
    fn register_and_query_struct() {
        let table = TypeTable::new();
        let handle = table.register_struct(
            "Windows.Foundation.Point",
            vec![
                ("X".into(), WinRTType::F32),
                ("Y".into(), WinRTType::F32),
            ],
        );
        assert_eq!(handle.size_of(), 8);
        assert_eq!(handle.align_of(), 4);
        assert!(table.contains("Windows.Foundation.Point"));

        let retrieved = table.get_struct_handle("Windows.Foundation.Point").unwrap();
        assert_eq!(retrieved, handle);
    }

    #[test]
    fn struct_deduplication() {
        let table = TypeTable::new();
        let h1 = table.register_struct(
            "Windows.Foundation.Point",
            vec![
                ("X".into(), WinRTType::F32),
                ("Y".into(), WinRTType::F32),
            ],
        );
        let h2 = table.register_struct(
            "Windows.Foundation.Point",
            vec![
                ("X".into(), WinRTType::F32),
                ("Y".into(), WinRTType::F32),
            ],
        );
        assert_eq!(h1, h2);
    }

    #[test]
    fn register_and_query_enum() {
        let table = TypeTable::new();
        table.register_enum(
            "Windows.Storage.FileAccessMode",
            WinRTType::I32,
            vec![("Read".into(), 0), ("ReadWrite".into(), 1)],
        );

        assert_eq!(
            table.get_enum_value("Windows.Storage.FileAccessMode", "Read"),
            Some(0)
        );
        assert_eq!(
            table.get_enum_value("Windows.Storage.FileAccessMode", "ReadWrite"),
            Some(1)
        );
        assert_eq!(
            table.get_enum_value("Windows.Storage.FileAccessMode", "Write"),
            None
        );
    }

    #[test]
    fn register_and_query_interface() {
        let iid = GUID::from_u128(0x9E365E57_48B2_4160_956F_C7385120BBFC);
        let table = TypeTable::new();
        table.register_interface(
            "Windows.Foundation.IUriRuntimeClass",
            iid,
            6,
            vec![
                (
                    "get_AbsoluteUri".into(),
                    MethodSignature::new().add_out(WinRTType::HString),
                ),
                (
                    "get_DisplayUri".into(),
                    MethodSignature::new().add_out(WinRTType::HString),
                ),
            ],
        );

        assert_eq!(
            table.get_iid("Windows.Foundation.IUriRuntimeClass"),
            Some(iid)
        );
        assert!(table.contains("Windows.Foundation.IUriRuntimeClass"));
    }

    #[test]
    fn register_and_query_runtime_class() {
        let default_iid = GUID::from_u128(0x9E365E57_48B2_4160_956F_C7385120BBFC);
        let table = TypeTable::new();
        table.register_runtime_class(
            "Windows.Foundation.Uri",
            default_iid,
            vec![InterfaceRef {
                name: "Windows.Foundation.IUriRuntimeClass".into(),
                role: InterfaceRole::Default,
            }],
            false,
        );

        assert_eq!(table.get_iid("Windows.Foundation.Uri"), Some(default_iid));
        let wt = table.get_winrt_type("Windows.Foundation.Uri").unwrap();
        match wt {
            WinRTType::RuntimeClass(name, iid) => {
                assert_eq!(name, "Windows.Foundation.Uri");
                assert_eq!(iid, default_iid);
            }
            _ => panic!("Expected RuntimeClass"),
        }
    }

    #[test]
    fn register_delegate() {
        let iid = GUID::from_u128(0x12345678_1234_1234_1234_123456789abc);
        let table = TypeTable::new();
        table.register_delegate("MyDelegate", iid);
        assert_eq!(table.get_iid("MyDelegate"), Some(iid));
    }

    #[test]
    fn get_winrt_type_for_all_kinds() {
        let table = TypeTable::new();

        // Struct
        table.register_struct(
            "Point",
            vec![("X".into(), WinRTType::F32), ("Y".into(), WinRTType::F32)],
        );
        assert!(matches!(
            table.get_winrt_type("Point"),
            Some(WinRTType::Struct(_))
        ));

        // Enum
        table.register_enum("MyEnum", WinRTType::I32, vec![("A".into(), 0)]);
        assert!(matches!(
            table.get_winrt_type("MyEnum"),
            Some(WinRTType::I32)
        ));

        // Interface
        let iid = GUID::from_u128(0x11111111_1111_1111_1111_111111111111);
        table.register_interface("IFoo", iid, 6, vec![]);
        assert!(matches!(
            table.get_winrt_type("IFoo"),
            Some(WinRTType::Interface(_))
        ));

        // Unknown
        assert!(table.get_winrt_type("DoesNotExist").is_none());
    }

    #[test]
    fn nested_struct_registration() {
        let table = TypeTable::new();

        // Register inner struct first
        let point_handle = table.register_struct(
            "Windows.Foundation.Point",
            vec![
                ("X".into(), WinRTType::F32),
                ("Y".into(), WinRTType::F32),
            ],
        );

        // Register outer struct using inner TypeHandle
        let ink_handle = table.register_struct(
            "Windows.UI.Input.Inking.InkTrailPoint",
            vec![
                ("Point".into(), WinRTType::Struct(point_handle)),
                ("Radius".into(), WinRTType::F32),
            ],
        );

        assert_eq!(ink_handle.size_of(), 12);
        assert_eq!(ink_handle.align_of(), 4);
        assert_eq!(ink_handle.field_count(), 2);
        assert_eq!(ink_handle.field_offset(0), 0);
        assert_eq!(ink_handle.field_offset(1), 8);
    }

    #[test]
    fn call_uri_via_type_table() {
        use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};
        use windows_core::{Interface, h};

        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };

        let table = TypeTable::new();

        // IUriRuntimeClassFactory (vtable index 6 = CreateUri)
        let factory_iid = GUID::from_u128(0x44A9796F_723E_4FDF_A218_033E75B0C084);
        table.register_interface(
            "Windows.Foundation.IUriRuntimeClassFactory",
            factory_iid,
            6,
            vec![(
                "CreateUri".into(),
                MethodSignature::new()
                    .add(WinRTType::HString)
                    .add_out(WinRTType::Object),
            )],
        );

        // IUriRuntimeClass (vtable index 6..= )
        let uri_iid = GUID::from_u128(0x9E365E57_48B2_4160_956F_C7385120BBFC);
        table.register_interface(
            "Windows.Foundation.IUriRuntimeClass",
            uri_iid,
            6,
            vec![
                (
                    "get_AbsoluteUri".into(),
                    MethodSignature::new().add_out(WinRTType::HString),
                ),
                (
                    "get_DisplayUri".into(),
                    MethodSignature::new().add_out(WinRTType::HString),
                ),
                (
                    "get_Domain".into(),
                    MethodSignature::new().add_out(WinRTType::HString),
                ),
                (
                    "get_Extension".into(),
                    MethodSignature::new().add_out(WinRTType::HString),
                ),
                (
                    "get_Fragment".into(),
                    MethodSignature::new().add_out(WinRTType::HString),
                ),
                (
                    "get_Host".into(),
                    MethodSignature::new().add_out(WinRTType::HString),
                ),
            ],
        );

        // Get factory
        let factory = unsafe {
            windows::Win32::System::WinRT::RoGetActivationFactory::<
                windows::Win32::System::WinRT::IActivationFactory,
            >(h!("Windows.Foundation.Uri"))
        }
        .unwrap();

        // QueryInterface for IUriRuntimeClassFactory
        let uri_factory: windows_core::IUnknown = factory.cast().unwrap();
        let mut factory_ptr = std::ptr::null_mut();
        unsafe {
            uri_factory
                .query(&factory_iid, &mut factory_ptr)
                .ok()
                .unwrap();
        }

        // Create Uri via TypeTable::call
        let results = table
            .call(
                "Windows.Foundation.IUriRuntimeClassFactory",
                "CreateUri",
                factory_ptr,
                &[WinRTValue::HString(windows_core::HSTRING::from(
                    "https://www.example.com/path?query=1#fragment",
                ))],
            )
            .unwrap();

        let uri_obj = results[0].as_object().unwrap();

        // Query IUriRuntimeClass
        let mut uri_ptr = std::ptr::null_mut();
        unsafe {
            uri_obj.query(&uri_iid, &mut uri_ptr).ok().unwrap();
        }

        // Call get_Host via TypeTable
        let host_results = table
            .call(
                "Windows.Foundation.IUriRuntimeClass",
                "get_Host",
                uri_ptr,
                &[],
            )
            .unwrap();

        let host = host_results[0].as_hstring().unwrap();
        assert_eq!(host.to_string(), "www.example.com");

        // Call get_AbsoluteUri via call_by_index
        let abs_results = table
            .call_by_index(
                "Windows.Foundation.IUriRuntimeClass",
                0, // get_AbsoluteUri is the first method
                uri_ptr,
                &[],
            )
            .unwrap();

        let abs_uri = abs_results[0].as_hstring().unwrap();
        assert_eq!(
            abs_uri.to_string(),
            "https://www.example.com/path?query=1#fragment"
        );
    }
}
