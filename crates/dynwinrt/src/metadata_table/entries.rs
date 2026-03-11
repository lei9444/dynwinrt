use windows_core::GUID;

use super::type_handle::TypeHandle;

// ===========================================================================
// Named entry types
// ===========================================================================

#[derive(Debug)]
pub enum TypeEntry {
    Interface(InterfaceEntry),
    RuntimeClass(RuntimeClassEntry),
    Struct(StructTypeEntry),
    Enum(EnumEntry),
    Delegate(DelegateEntry),
}

/// A WinRT interface. Methods are stored in the global methods arena;
/// `method_indices` holds their positions there.
#[derive(Debug)]
pub struct InterfaceEntry {
    pub name: String,
    pub iid: GUID,
    pub method_names: Vec<String>,
    pub method_indices: Vec<u32>,
}

#[derive(Debug)]
pub struct RuntimeClassEntry {
    pub name: String,
    pub default_iid: GUID,
    pub interfaces: Vec<InterfaceRef>,
    pub has_default_constructor: bool,
}

#[derive(Debug, Clone)]
pub struct InterfaceRef {
    pub name: String,
    pub role: InterfaceRole,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InterfaceRole {
    Default,
    Factory,
    Static,
    Other,
}

#[derive(Debug)]
pub struct StructTypeEntry {
    pub name: String,
    pub type_handle: TypeHandle,
    pub field_names: Vec<String>,
    pub field_types: Vec<TypeHandle>,
}

#[derive(Debug)]
pub struct EnumEntry {
    pub name: String,
    pub underlying: TypeHandle,
    pub members: Vec<(String, i32)>,
}

#[derive(Debug)]
pub struct DelegateEntry {
    pub name: String,
    pub iid: GUID,
}
