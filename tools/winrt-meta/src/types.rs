/// Describes a WinRT type as extracted from WinMD metadata.
#[derive(Debug, Clone, PartialEq)]
pub enum TypeMeta {
    // Primitives
    Bool,
    I8,
    U8,
    I16,
    U16,
    I32,
    U32,
    I64,
    U64,
    F32,
    F64,
    Char16,
    String, // HSTRING
    Guid,

    // Reference types
    Object, // IInspectable / unknown type
    Interface {
        namespace: String,
        name: String,
        iid: String,
    },
    RuntimeClass {
        namespace: String,
        name: String,
        default_iid: String,
    },
    Delegate {
        namespace: String,
        name: String,
        iid: String,
    },

    // Async patterns
    AsyncAction,
    AsyncActionWithProgress(Box<TypeMeta>),
    AsyncOperation(Box<TypeMeta>),
    AsyncOperationWithProgress(Box<TypeMeta>, Box<TypeMeta>),

    // Parameterized interface instantiation: e.g. IVector<String>, IMap<String, Object>
    Parameterized {
        namespace: String,
        name: String,
        piid: String,
        args: Vec<TypeMeta>,
    },

    // Composite
    Array(Box<TypeMeta>),
    Struct {
        namespace: String,
        name: String,
        fields: Vec<FieldMeta>,
    },
    Enum {
        namespace: String,
        name: String,
        underlying: Box<TypeMeta>,
        members: Vec<EnumMember>,
    },
}

#[derive(Debug, Clone, PartialEq)]
pub struct FieldMeta {
    pub name: String,
    pub typ: TypeMeta,
}

#[derive(Debug, Clone, PartialEq)]
pub struct EnumMember {
    pub name: String,
    pub value: i32,
}

impl TypeMeta {
    /// Returns true if this type represents an async operation.
    pub fn is_async(&self) -> bool {
        matches!(
            self,
            TypeMeta::AsyncAction
                | TypeMeta::AsyncActionWithProgress(_)
                | TypeMeta::AsyncOperation(_)
                | TypeMeta::AsyncOperationWithProgress(_, _)
        )
    }

    /// For async types, return the result type (if any).
    pub fn async_result_type(&self) -> Option<&TypeMeta> {
        match self {
            TypeMeta::AsyncOperation(t) | TypeMeta::AsyncOperationWithProgress(t, _) => Some(t),
            _ => None,
        }
    }
}
