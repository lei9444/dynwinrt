use crate::metadata_table::TypeKind;
use crate::abi::AbiType;

#[derive(Debug)]
pub enum Error {
    ExpectObjectTypeError(TypeKind),
    InvalidType(TypeKind, TypeKind),
    InvalidNestedOutType(TypeKind),
    InvalidTypeAbiToWinRT(TypeKind, AbiType),
    WindowsError(windows_core::Error),
    TypeNotFound(String),
    NotAnInterface(String),
    MethodNotFound(String, String),
}

impl Error {
    pub fn expect_object_type(actual: TypeKind) -> Self {
        Error::ExpectObjectTypeError(actual)
    }

    pub fn message(&self) -> String {
        match self {
            Error::ExpectObjectTypeError(actual) => {
                format!("Expected object type, found {:?}", actual)
            }
            Error::InvalidType(expected, actual) => {
                format!("Invalid type: expected {:?}, found {:?}", expected, actual)
            }
            Error::InvalidNestedOutType(actual) => {
                format!("Invalid nested out type: found {:?}", actual)
            }
            Error::InvalidTypeAbiToWinRT(expected, actual) => {
                format!(
                    "Invalid type ABI to WinRT: expected {:?}, found {:?}",
                    expected, actual
                )
            }
            Error::WindowsError(err) => format!("Windows error: {}", err),
            Error::TypeNotFound(name) => format!("Type not found: {}", name),
            Error::NotAnInterface(name) => format!("Not an interface: {}", name),
            Error::MethodNotFound(iface, method) => {
                format!("Method '{}' not found on interface '{}'", method, iface)
            }
        }
    }
}

impl From<windows::core::Error> for Error {
    fn from(value: windows::core::Error) -> Self {
        Self::WindowsError(value)
    }
}

pub type Result<T> = core::result::Result<T, Error>;
