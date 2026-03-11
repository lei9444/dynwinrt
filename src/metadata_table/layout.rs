use std::alloc::Layout;

use super::MetadataTable;
use super::type_kind::TypeKind;

// ===========================================================================
// Internal indexed storage types
// ===========================================================================

pub(crate) struct RuntimeClassData {
    pub name: String,
    pub default_iid: windows_core::GUID,
}

pub(crate) struct ParameterizedData {
    pub generic_def: TypeKind,
    pub args: Vec<TypeKind>,
}

pub(super) struct StructEntry {
    /// WinRT full name (e.g. "Windows.Web.Http.HttpProgress") for IID signature computation.
    /// None for anonymous structs created via define_struct().
    pub(super) name: Option<String>,
    pub(super) field_kinds: Vec<TypeKind>,
    pub(super) field_offsets: Vec<usize>,
    pub(super) layout: Layout,
}

// ===========================================================================
// Layout engine + indexed data access on MetadataTable
// ===========================================================================

impl MetadataTable {
    // -----------------------------------------------------------------------
    // Internal: index registration
    // -----------------------------------------------------------------------

    pub(crate) fn register_inner_type(&self, kind: TypeKind) -> u32 {
        let mut inner = self.inner_types.write().unwrap();
        let idx = inner.len() as u32;
        inner.push(kind);
        idx
    }

    pub(crate) fn register_inner_type_pair(&self, a: TypeKind, b: TypeKind) -> u32 {
        let mut pairs = self.inner_type_pairs.write().unwrap();
        let idx = pairs.len() as u32;
        pairs.push((a, b));
        idx
    }

    // -----------------------------------------------------------------------
    // Internal: indexed data access
    // -----------------------------------------------------------------------

    pub(crate) fn get_inner_type(&self, idx: u32) -> TypeKind {
        self.inner_types.read().unwrap()[idx as usize]
    }

    pub(crate) fn get_inner_type_pair(&self, idx: u32) -> (TypeKind, TypeKind) {
        self.inner_type_pairs.read().unwrap()[idx as usize]
    }

    pub(crate) fn get_runtime_class(&self, idx: u32) -> (String, windows_core::GUID) {
        let rcs = self.runtime_classes.read().unwrap();
        let rc = &rcs[idx as usize];
        (rc.name.clone(), rc.default_iid)
    }

    pub(crate) fn get_parameterized(&self, idx: u32) -> (TypeKind, Vec<TypeKind>) {
        let pts = self.parameterized_types.read().unwrap();
        let p = &pts[idx as usize];
        (p.generic_def, p.args.clone())
    }

    // -----------------------------------------------------------------------
    // Layout engine
    // -----------------------------------------------------------------------

    pub(crate) fn size_of_kind(&self, kind: TypeKind) -> usize {
        if let Some(s) = kind.primitive_size() {
            return s;
        }
        match kind {
            TypeKind::Struct(id) => self.structs.read().unwrap()[id as usize].layout.size(),
            TypeKind::HString | TypeKind::Object
            | TypeKind::Interface(_) | TypeKind::Delegate(_)
            | TypeKind::RuntimeClass(_) | TypeKind::Parameterized(_)
            | TypeKind::IAsyncAction | TypeKind::IAsyncActionWithProgress(_)
            | TypeKind::IAsyncOperation(_) | TypeKind::IAsyncOperationWithProgress(_)
            | TypeKind::OutValue(_) | TypeKind::ArrayOfIUnknown => {
                std::mem::size_of::<*mut std::ffi::c_void>()
            }
            _ => panic!("size_of_kind not supported for {:?}", kind),
        }
    }

    pub(crate) fn align_of_kind(&self, kind: TypeKind) -> usize {
        if let Some(a) = kind.primitive_align() {
            return a;
        }
        match kind {
            TypeKind::Struct(id) => self.structs.read().unwrap()[id as usize].layout.align(),
            TypeKind::HString | TypeKind::Object
            | TypeKind::Interface(_) | TypeKind::Delegate(_)
            | TypeKind::RuntimeClass(_) | TypeKind::Parameterized(_)
            | TypeKind::IAsyncAction | TypeKind::IAsyncActionWithProgress(_)
            | TypeKind::IAsyncOperation(_) | TypeKind::IAsyncOperationWithProgress(_)
            | TypeKind::OutValue(_) | TypeKind::ArrayOfIUnknown => {
                std::mem::align_of::<*mut std::ffi::c_void>()
            }
            _ => panic!("align_of_kind not supported for {:?}", kind),
        }
    }

    pub(crate) fn layout_of_kind(&self, kind: TypeKind) -> Layout {
        let size = self.size_of_kind(kind);
        let align = self.align_of_kind(kind);
        Layout::from_size_align(size, align).unwrap()
    }

    pub(crate) fn field_count_kind(&self, kind: TypeKind) -> usize {
        match kind {
            TypeKind::Struct(id) => self.structs.read().unwrap()[id as usize].field_kinds.len(),
            _ => panic!("field_count: type {:?} has no fields", kind),
        }
    }

    pub(crate) fn field_offset_kind(&self, kind: TypeKind, index: usize) -> usize {
        match kind {
            TypeKind::Struct(id) => self.structs.read().unwrap()[id as usize].field_offsets[index],
            _ => panic!("field_offset: type {:?} has no fields", kind),
        }
    }

    pub(crate) fn field_kind(&self, kind: TypeKind, index: usize) -> TypeKind {
        match kind {
            TypeKind::Struct(id) => self.structs.read().unwrap()[id as usize].field_kinds[index],
            _ => panic!("field_kind: type {:?} has no fields", kind),
        }
    }

    pub(crate) fn libffi_type_kind(&self, kind: TypeKind) -> libffi::middle::Type {
        if let Some(t) = kind.primitive_libffi_type() {
            return t;
        }
        match kind {
            TypeKind::Struct(id) => {
                let structs = self.structs.read().unwrap();
                let field_types: Vec<libffi::middle::Type> = structs[id as usize]
                    .field_kinds
                    .iter()
                    .map(|f| self.libffi_type_kind(*f))
                    .collect();
                libffi::middle::Type::structure(field_types)
            }
            _ => panic!("libffi_type_kind: unsupported for {:?}", kind),
        }
    }

    pub(super) fn compute_layout(&self, fields: &[TypeKind]) -> (Vec<usize>, Layout) {
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
