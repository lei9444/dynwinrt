use std::collections::HashSet;

use windows_metadata::{HasAttributes, reader};

use crate::types::{EnumMember, TypeMeta};

/// Direction of a method parameter at the ABI level.
#[derive(Debug, Clone, PartialEq)]
pub enum ParamDirection {
    In,
    Out,
    /// FillArray: caller allocates buffer, callee fills it.
    OutFill,
}

/// A single method parameter.
#[derive(Debug, Clone)]
pub struct ParamMeta {
    pub name: String,
    pub typ: TypeMeta,
    pub direction: ParamDirection,
}

/// A method on a WinRT interface.
#[derive(Debug, Clone)]
pub struct MethodMeta {
    pub name: String,
    pub vtable_index: usize,
    pub params: Vec<ParamMeta>,
    pub return_type: Option<TypeMeta>,
    pub is_property_getter: bool,
    pub is_property_setter: bool,
    pub is_event_add: bool,
    pub is_event_remove: bool,
}

/// A WinRT interface with its methods.
#[derive(Debug, Clone)]
pub struct InterfaceMeta {
    pub name: String,
    pub namespace: String,
    pub iid: String,
    pub methods: Vec<MethodMeta>,
}

/// How an interface relates to a RuntimeClass.
#[derive(Debug, Clone, PartialEq)]
pub enum InterfaceRole {
    Default,
    Factory,
    Static,
    Other,
}

/// A WinRT RuntimeClass with all its interfaces.
#[derive(Debug, Clone)]
pub struct ClassMeta {
    pub name: String,
    pub namespace: String,
    pub full_name: String,
    pub default_interface: Option<InterfaceMeta>,
    pub factory_interfaces: Vec<InterfaceMeta>,
    pub static_interfaces: Vec<InterfaceMeta>,
    pub has_default_constructor: bool,
}

/// Parse a WinMD file and extract metadata for a single RuntimeClass.
/// Accepts multiple winmd paths separated by ';'.
pub fn parse_class(winmd_paths: &str, namespace: &str, name: &str) -> Option<ClassMeta> {
    let index = load_index(winmd_paths)?;
    parse_class_from_index(&index, namespace, name)
}

/// Parse all RuntimeClasses in a given namespace.
pub fn parse_namespace(winmd_paths: &str, namespace: &str) -> Vec<ClassMeta> {
    let index = match load_index(winmd_paths) {
        Some(idx) => idx,
        None => return Vec::new(),
    };

    let mut classes = Vec::new();
    for def in index.all() {
        if def.namespace() != namespace {
            continue;
        }
        let extends = match def.extends() {
            Some(e) => e,
            None => continue,
        };
        if extends.namespace() != "System" || extends.name() != "Object" {
            continue;
        }

        if let Some(class) = parse_class_from_index(&index, namespace, def.name()) {
            classes.push(class);
        }
    }
    classes
}

/// Parse public (non-exclusive) interfaces in a namespace.
/// Exclusive interfaces (prefixed with I and paired with a RuntimeClass) are skipped
/// since they are implementation details. We only generate public-facing interfaces.
pub fn parse_interfaces(winmd_paths: &str, namespace: &str) -> Vec<InterfaceMeta> {
    let index = match load_index(winmd_paths) {
        Some(idx) => idx,
        None => return Vec::new(),
    };

    let mut interfaces = Vec::new();
    for def in index.all() {
        if def.namespace() != namespace {
            continue;
        }
        // Interfaces have no extends (or extend nothing)
        if def.extends().is_some() {
            continue;
        }
        // Skip generic interface definitions (they have generic params)
        if def.generic_params().next().is_some() {
            continue;
        }
        // Check it's actually an interface by looking for GuidAttribute
        let iid = extract_iid(&def);
        if iid.is_empty() {
            continue;
        }
        // Skip exclusive interfaces (marked with ExclusiveTo attribute)
        if def.has_attribute("ExclusiveToAttribute") {
            continue;
        }
        if let Some(iface) = parse_interface(&index, namespace, def.name()) {
            interfaces.push(iface);
        }
    }
    interfaces
}

/// Parse enums in a namespace.
pub fn parse_enums(winmd_paths: &str, namespace: &str) -> Vec<TypeMeta> {
    let index = match load_index(winmd_paths) {
        Some(idx) => idx,
        None => return Vec::new(),
    };

    let mut enums = Vec::new();
    for def in index.all() {
        if def.namespace() != namespace {
            continue;
        }
        if let Some(extends) = def.extends() {
            if extends.namespace() == "System" && extends.name() == "Enum" {
                enums.push(parse_enum_def(&def));
            }
        }
    }
    enums
}

/// Collect all type references from a class (both same-namespace and cross-namespace).
/// Excludes the class itself.
pub fn collect_imports(class: &ClassMeta) -> HashSet<(String, String)> {
    let mut imports: HashSet<(String, String)> = HashSet::new();
    let class_name = &class.name;

    fn visit_type(typ: &TypeMeta, class_name: &str, imports: &mut HashSet<(String, String)>) {
        let mut named = Vec::new();
        let mut _param = Vec::new();
        visit_type_refs(typ, &mut named, &mut _param);
        for (ns, name, _kind) in named {
            if name != class_name {
                imports.insert((ns, name));
            }
        }
    }

    fn visit_methods(methods: &[MethodMeta], class_name: &str, imports: &mut HashSet<(String, String)>) {
        for m in methods {
            for p in &m.params {
                visit_type(&p.typ, class_name, imports);
            }
            if let Some(ref rt) = m.return_type {
                visit_type(rt, class_name, imports);
            }
        }
    }

    if let Some(ref iface) = class.default_interface {
        visit_methods(&iface.methods, class_name, &mut imports);
    }
    for iface in &class.factory_interfaces {
        visit_methods(&iface.methods, class_name, &mut imports);
    }
    for iface in &class.static_interfaces {
        visit_methods(&iface.methods, class_name, &mut imports);
    }

    imports
}

/// Resolved dependency types that need to be generated.
pub struct ResolvedDeps {
    pub classes: Vec<ClassMeta>,
    pub interfaces: Vec<InterfaceMeta>,
    pub enums: Vec<TypeMeta>,
}

/// Resolve all referenced types that don't have generated files yet.
/// Uses fixpoint iteration to recursively discover transitive dependencies.
pub fn resolve_dependencies(
    winmd_paths: &str,
    classes: &[ClassMeta],
    existing_interfaces: &[InterfaceMeta],
    existing_enums: &[TypeMeta],
) -> ResolvedDeps {
    let index = match load_index(winmd_paths) {
        Some(idx) => idx,
        None => return ResolvedDeps { classes: vec![], interfaces: vec![], enums: vec![] },
    };

    // Track all known type names (already generated or discovered)
    let mut known: HashSet<String> = HashSet::new();
    for c in classes { known.insert(c.name.clone()); }
    for i in existing_interfaces { known.insert(i.name.clone()); }
    for e in existing_enums {
        if let TypeMeta::Enum { name, .. } = e { known.insert(name.clone()); }
    }

    let mut dep_classes: Vec<ClassMeta> = Vec::new();
    let mut dep_interfaces: Vec<InterfaceMeta> = Vec::new();
    let mut dep_enums: Vec<TypeMeta> = Vec::new();

    // Seed the worklist from initial types
    let mut worklist: Vec<(String, String, &'static str)> = Vec::new();
    let mut param_worklist: Vec<TypeMeta> = Vec::new();
    collect_all_refs_from_classes(classes, &known, &mut worklist, &mut param_worklist);
    collect_all_refs_from_interfaces(existing_interfaces, &known, &mut worklist, &mut param_worklist);

    // Fixpoint: keep resolving until no new types are discovered
    loop {
        let has_work = !worklist.is_empty() || !param_worklist.is_empty();
        if !has_work { break; }

        let batch: Vec<_> = worklist.drain(..).collect();
        let param_batch: Vec<_> = param_worklist.drain(..).collect();
        let mut new_classes = Vec::new();
        let mut new_interfaces = Vec::new();

        for (ns, name, kind) in &batch {
            if known.contains(name) { continue; }
            known.insert(name.clone());

            match *kind {
                "interface" => {
                    if let Some(iface) = parse_interface(&index, ns, name) {
                        new_interfaces.push(iface);
                    }
                }
                "class" => {
                    if let Some(class) = parse_class_from_index(&index, ns, name) {
                        new_classes.push(class);
                    }
                }
                "enum" => {
                    if let Some(def) = index.get(ns, name).next() {
                        dep_enums.push(parse_enum_def(&def));
                    }
                }
                _ => {}
            }
        }

        // Resolve parameterized interfaces (e.g. IVector<String>)
        for param_type in &param_batch {
            if let TypeMeta::Parameterized { namespace, name, piid, args } = param_type {
                let concrete_name = make_parameterized_name(name, args);
                if known.contains(&concrete_name) { continue; }
                known.insert(concrete_name.clone());

                if let Some(iface) = parse_parameterized_interface(
                    &index, namespace, name, &concrete_name, piid, args,
                ) {
                    new_interfaces.push(iface);
                }
            }
        }

        // Discover new references from the newly resolved types
        collect_all_refs_from_classes(&new_classes, &known, &mut worklist, &mut param_worklist);
        collect_all_refs_from_interfaces(&new_interfaces, &known, &mut worklist, &mut param_worklist);

        dep_classes.extend(new_classes);
        dep_interfaces.extend(new_interfaces);
    }

    ResolvedDeps { classes: dep_classes, interfaces: dep_interfaces, enums: dep_enums }
}

/// Visit a TypeMeta tree and collect both named type references and parameterized types.
fn visit_type_refs(
    typ: &TypeMeta,
    named: &mut Vec<(String, String, &'static str)>,
    parameterized: &mut Vec<TypeMeta>,
) {
    match typ {
        TypeMeta::Interface { namespace, name, .. } => {
            named.push((namespace.clone(), name.clone(), "interface"));
        }
        TypeMeta::RuntimeClass { namespace, name, .. } => {
            named.push((namespace.clone(), name.clone(), "class"));
        }
        TypeMeta::Enum { namespace, name, .. } => {
            named.push((namespace.clone(), name.clone(), "enum"));
        }
        TypeMeta::AsyncOperation(inner) | TypeMeta::AsyncActionWithProgress(inner) => {
            visit_type_refs(inner, named, parameterized);
        }
        TypeMeta::AsyncOperationWithProgress(r, p) => {
            visit_type_refs(r, named, parameterized);
            visit_type_refs(p, named, parameterized);
        }
        TypeMeta::Struct { fields, .. } => {
            for f in fields {
                visit_type_refs(&f.typ, named, parameterized);
            }
        }
        TypeMeta::Array(inner) => {
            visit_type_refs(inner, named, parameterized);
        }
        TypeMeta::Parameterized { args, .. } => {
            parameterized.push(typ.clone());
            for arg in args {
                visit_type_refs(arg, named, parameterized);
            }
        }
        _ => {}
    }
}

/// Collect all type references from methods: both named and parameterized.
fn collect_all_refs_from_methods(
    methods: &[MethodMeta],
    known: &HashSet<String>,
    named_out: &mut Vec<(String, String, &'static str)>,
    param_out: &mut Vec<TypeMeta>,
) {
    let mut named = Vec::new();
    let mut parameterized = Vec::new();
    for m in methods {
        for p in &m.params {
            visit_type_refs(&p.typ, &mut named, &mut parameterized);
        }
        if let Some(ref rt) = m.return_type {
            visit_type_refs(rt, &mut named, &mut parameterized);
        }
    }
    for r in named {
        if !known.contains(&r.1) {
            named_out.push(r);
        }
    }
    for r in parameterized {
        if let TypeMeta::Parameterized { name, args, .. } = &r {
            let concrete = make_parameterized_name(name, args);
            if !known.contains(&concrete) {
                param_out.push(r);
            }
        }
    }
}

/// Generate a concrete name for a parameterized interface, e.g. "IVector_String", "IMap_String_Object".
pub fn make_parameterized_name(generic_name: &str, args: &[TypeMeta]) -> String {
    let base = generic_name.split('`').next().unwrap_or(generic_name);
    let arg_names: Vec<String> = args.iter().map(|a| type_meta_short_name(a)).collect();
    format!("{}_{}", base, arg_names.join("_"))
}

fn type_meta_short_name(typ: &TypeMeta) -> String {
    match typ {
        TypeMeta::Bool => "Boolean".to_string(),
        TypeMeta::I8 => "Int8".to_string(),
        TypeMeta::U8 => "UInt8".to_string(),
        TypeMeta::I16 => "Int16".to_string(),
        TypeMeta::U16 => "UInt16".to_string(),
        TypeMeta::I32 => "Int32".to_string(),
        TypeMeta::U32 => "UInt32".to_string(),
        TypeMeta::I64 => "Int64".to_string(),
        TypeMeta::U64 => "UInt64".to_string(),
        TypeMeta::F32 => "Single".to_string(),
        TypeMeta::F64 => "Double".to_string(),
        TypeMeta::String => "String".to_string(),
        TypeMeta::Char16 => "Char16".to_string(),
        TypeMeta::Guid => "Guid".to_string(),
        TypeMeta::Object => "Object".to_string(),
        TypeMeta::RuntimeClass { name, .. }
        | TypeMeta::Interface { name, .. }
        | TypeMeta::Enum { name, .. } => name.clone(),
        TypeMeta::Parameterized { name, args, .. } => make_parameterized_name(name, args),
        _ => "Unknown".to_string(),
    }
}

/// Collect all refs from a list of classes (iterates all interface methods).
fn collect_all_refs_from_classes(
    classes: &[ClassMeta],
    known: &HashSet<String>,
    named_out: &mut Vec<(String, String, &'static str)>,
    param_out: &mut Vec<TypeMeta>,
) {
    for c in classes {
        if let Some(ref iface) = c.default_interface {
            collect_all_refs_from_methods(&iface.methods, known, named_out, param_out);
        }
        for iface in &c.factory_interfaces {
            collect_all_refs_from_methods(&iface.methods, known, named_out, param_out);
        }
        for iface in &c.static_interfaces {
            collect_all_refs_from_methods(&iface.methods, known, named_out, param_out);
        }
    }
}

/// Collect all refs from a list of standalone interfaces.
fn collect_all_refs_from_interfaces(
    interfaces: &[InterfaceMeta],
    known: &HashSet<String>,
    named_out: &mut Vec<(String, String, &'static str)>,
    param_out: &mut Vec<TypeMeta>,
) {
    for i in interfaces {
        collect_all_refs_from_methods(&i.methods, known, named_out, param_out);
    }
}

// --- Internal helpers ---

/// Well-known PIIDs for generic interfaces whose GuidAttribute can't be read via extract_iid.
fn well_known_piid(namespace: &str, name: &str) -> Option<&'static str> {
    if namespace == "Windows.Foundation" {
        match name {
            "IReference`1" => Some("61c17706-2d65-11e0-9ae8-d48564015472"),
            _ => None,
        }
    } else if namespace == "Windows.Foundation.Collections" {
        match name {
            "IIterable`1" => Some("faa585ea-6214-4217-afda-7f46de5869b3"),
            "IIterator`1" => Some("6a79e863-4300-459a-9966-cbb660963ee1"),
            "IVectorView`1" => Some("bbe1fa4c-b0e3-4583-baef-1f1b2e483e56"),
            "IVector`1" => Some("913337e9-11a1-4345-a3a2-4e7f956e222d"),
            "IMapView`2" => Some("e480ce40-a338-4ada-adcf-272272e48cb9"),
            "IMap`2" => Some("3c2925fe-8519-45c1-aa79-197b6718c1c1"),
            "IKeyValuePair`2" => Some("02b51929-c1c4-4a7e-8940-0312b5c18500"),
            "IObservableVector`1" => Some("5917eb53-50b4-4a0d-b309-65862b3f1dbc"),
            "IObservableMap`2" => Some("65df2bf5-bf39-41b5-aebc-5a9d865e472b"),
            _ => None,
        }
    } else {
        None
    }
}

fn load_index(winmd_paths: &str) -> Option<reader::Index> {
    let paths: Vec<&str> = winmd_paths.split(';').filter(|s| !s.is_empty()).collect();
    if paths.len() == 1 {
        reader::Index::read(paths[0])
    } else {
        // Load multiple winmd files
        let mut files = Vec::new();
        for path in &paths {
            if let Some(f) = reader::File::read(path) {
                files.push(f);
            }
        }
        if files.is_empty() {
            None
        } else {
            Some(reader::Index::new(files))
        }
    }
}

fn parse_class_from_index(index: &reader::Index, namespace: &str, name: &str) -> Option<ClassMeta> {
    let def = index.get(namespace, name).next()?;
    let full_name = format!("{}.{}", namespace, name);

    let mut default_interface = None;
    let mut factory_interfaces = Vec::new();
    let mut static_interfaces = Vec::new();
    let mut has_default_constructor = false;

    // 1. Find default interface from InterfaceImpl attributes
    for iface_impl in def.interface_impls() {
        let iface_ty = iface_impl.interface(&[]);
        let (iface_ns, iface_name) = match &iface_ty {
            windows_metadata::Type::Name(tn) => (tn.namespace.clone(), tn.name.clone()),
            _ => continue,
        };

        if iface_impl.has_attribute("DefaultAttribute") {
            if let Some(iface_meta) = parse_interface(index, &iface_ns, &iface_name) {
                default_interface = Some(iface_meta);
            }
        }
    }

    // 2. Find factory/static/default-constructor from class-level attributes
    for attr in def.attributes() {
        let attr_name = attr.ctor().parent().name().to_string();
        let values = attr.value();

        if attr_name == "ActivatableAttribute" {
            match values.first() {
                Some((_, windows_metadata::Value::Utf8(iface_full_name))) => {
                    // Factory interface specified
                    if let Some((ns, n)) = split_full_name(iface_full_name) {
                        if let Some(iface_meta) = parse_interface(index, ns, n) {
                            factory_interfaces.push(iface_meta);
                        }
                    }
                }
                Some((_, windows_metadata::Value::U32(_))) | Some((_, windows_metadata::Value::I32(_))) => {
                    // No factory interface — this is a default (parameterless) constructor
                    has_default_constructor = true;
                }
                _ => {
                    has_default_constructor = true;
                }
            }
        } else if attr_name == "StaticAttribute" {
            if let Some((_, windows_metadata::Value::Utf8(iface_full_name))) = values.first() {
                if let Some((ns, n)) = split_full_name(iface_full_name) {
                    if let Some(iface_meta) = parse_interface(index, ns, n) {
                        static_interfaces.push(iface_meta);
                    }
                }
            }
        }
    }

    Some(ClassMeta {
        name: name.to_string(),
        namespace: namespace.to_string(),
        full_name,
        default_interface,
        factory_interfaces,
        static_interfaces,
        has_default_constructor,
    })
}

fn split_full_name(full_name: &str) -> Option<(&str, &str)> {
    let dot_pos = full_name.rfind('.')?;
    Some((&full_name[..dot_pos], &full_name[dot_pos + 1..]))
}

fn parse_interface(
    index: &reader::Index,
    namespace: &str,
    name: &str,
) -> Option<InterfaceMeta> {
    let def = index.get(namespace, name).next()?;
    let iid = extract_iid(&def);
    parse_interface_methods(index, &def, name, namespace, &iid, &[])
}

/// Parse a parameterized interface definition (e.g. IVector`1) from winmd,
/// substituting generic type parameters with concrete types.
/// Returns an InterfaceMeta with a mangled name like "IVector_String".
fn parse_parameterized_interface(
    index: &reader::Index,
    namespace: &str,
    generic_name: &str,
    concrete_name: &str,
    piid: &str,
    generic_args: &[TypeMeta],
) -> Option<InterfaceMeta> {
    let trimmed_name = generic_name.split('`').next().unwrap_or(generic_name);
    let def = index.get(namespace, trimmed_name).next()?;
    parse_interface_methods(index, &def, concrete_name, namespace, piid, generic_args)
}

/// Core interface parsing: extract methods from a TypeDef, optionally substituting generics.
fn parse_interface_methods(
    index: &reader::Index,
    def: &reader::TypeDef,
    output_name: &str,
    namespace: &str,
    iid: &str,
    generic_args: &[TypeMeta],
) -> Option<InterfaceMeta> {
    let winmd_generics: Vec<windows_metadata::Type> =
        generic_args.iter().map(type_meta_to_winmd_type).collect();

    let mut methods = Vec::new();
    for (i, method) in def.methods().enumerate() {
        let vtable_index = 6 + i;
        let sig = method.signature(&winmd_generics);

        let overload_name = method.find_attribute("OverloadAttribute").and_then(|a| {
            a.value().into_iter().next().and_then(|(_, v)| match v {
                windows_metadata::Value::Utf8(s) => Some(s),
                _ => None,
            })
        });
        let method_name = overload_name.unwrap_or_else(|| method.name().to_string());

        let mut params = Vec::new();
        let param_defs: Vec<_> = method.params().collect();
        for (j, param_def) in param_defs.iter().enumerate() {
            if j < sig.types.len() {
                let typ = map_winmd_type_with_generics(&sig.types[j], index, generic_args);
                let is_out = param_def.flags().contains(windows_metadata::ParamAttributes::Out);
                let direction = if is_out {
                    if matches!(sig.types[j], windows_metadata::Type::Array(_)) {
                        // [out] Array = FillArray (caller allocates buffer)
                        ParamDirection::OutFill
                    } else {
                        // [out] ArrayRef or scalar = ReceiveArray / regular out
                        ParamDirection::Out
                    }
                } else {
                    ParamDirection::In
                };
                params.push(ParamMeta {
                    name: param_def.name().to_string(),
                    typ,
                    direction,
                });
            }
        }

        let return_type = if sig.return_type == windows_metadata::Type::Void {
            None
        } else {
            Some(map_winmd_type_with_generics(&sig.return_type, index, generic_args))
        };

        methods.push(MethodMeta {
            name: method_name.clone(),
            vtable_index,
            params,
            return_type,
            is_property_getter: method_name.starts_with("get_"),
            is_property_setter: method_name.starts_with("put_"),
            is_event_add: method_name.starts_with("add_"),
            is_event_remove: method_name.starts_with("remove_"),
        });
    }

    Some(InterfaceMeta {
        name: output_name.to_string(),
        namespace: namespace.to_string(),
        iid: iid.to_string(),
        methods,
    })
}

/// Convert TypeMeta back to windows_metadata::Type (for passing to method.signature()).
fn type_meta_to_winmd_type(typ: &TypeMeta) -> windows_metadata::Type {
    match typ {
        TypeMeta::Bool => windows_metadata::Type::Bool,
        TypeMeta::I8 => windows_metadata::Type::I8,
        TypeMeta::U8 => windows_metadata::Type::U8,
        TypeMeta::I16 => windows_metadata::Type::I16,
        TypeMeta::U16 => windows_metadata::Type::U16,
        TypeMeta::I32 => windows_metadata::Type::I32,
        TypeMeta::U32 => windows_metadata::Type::U32,
        TypeMeta::I64 => windows_metadata::Type::I64,
        TypeMeta::U64 => windows_metadata::Type::U64,
        TypeMeta::F32 => windows_metadata::Type::F32,
        TypeMeta::F64 => windows_metadata::Type::F64,
        TypeMeta::String => windows_metadata::Type::String,
        TypeMeta::Char16 => windows_metadata::Type::Char,
        TypeMeta::Guid => windows_metadata::Type::named("System", "Guid"),
        TypeMeta::Object => windows_metadata::Type::Object,
        TypeMeta::RuntimeClass { namespace, name, .. }
        | TypeMeta::Interface { namespace, name, .. }
        | TypeMeta::Enum { namespace, name, .. } => windows_metadata::Type::named(namespace, name),
        _ => windows_metadata::Type::Object,
    }
}

fn extract_iid(def: &reader::TypeDef) -> String {
    if let Some(attr) = def.find_attribute("GuidAttribute") {
        let args: Vec<(String, windows_metadata::Value)> = attr.value();
        if args.len() >= 11 {
            let a = extract_u32(&args[0].1);
            let b = extract_u16(&args[1].1);
            let c = extract_u16(&args[2].1);
            let d = extract_u8(&args[3].1);
            let e = extract_u8(&args[4].1);
            let f = extract_u8(&args[5].1);
            let g = extract_u8(&args[6].1);
            let h = extract_u8(&args[7].1);
            let i = extract_u8(&args[8].1);
            let j = extract_u8(&args[9].1);
            let k = extract_u8(&args[10].1);
            return format!(
                "{:08x}-{:04x}-{:04x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}",
                a, b, c, d, e, f, g, h, i, j, k
            );
        }
    }
    String::new()
}

fn extract_u32(val: &windows_metadata::Value) -> u32 {
    match val {
        windows_metadata::Value::U32(v) => *v,
        _ => 0,
    }
}

fn extract_u16(val: &windows_metadata::Value) -> u16 {
    match val {
        windows_metadata::Value::U16(v) => *v,
        _ => 0,
    }
}

fn extract_u8(val: &windows_metadata::Value) -> u8 {
    match val {
        windows_metadata::Value::U8(v) => *v,
        _ => 0,
    }
}

fn find_default_interface_iid(def: &reader::TypeDef, index: &reader::Index) -> String {
    for iface_impl in def.interface_impls() {
        if !iface_impl.has_attribute("DefaultAttribute") {
            continue;
        }
        let iface_ty = iface_impl.interface(&[]);
        if let windows_metadata::Type::Name(tn) = &iface_ty {
            if let Some(iface_def) = index.get(&tn.namespace, &tn.name).next() {
                let iid = extract_iid(&iface_def);
                if !iid.is_empty() {
                    return iid;
                }
            }
        }
    }
    String::new()
}

fn parse_enum_def(def: &reader::TypeDef) -> TypeMeta {
    let mut members = Vec::new();
    for field in def.fields() {
        let name = field.name().to_string();
        if name == "value__" {
            continue; // Skip the underlying value field
        }
        // Enum fields have constant values
        if let Some(constant) = field.constant() {
            let value = match constant.value() {
                windows_metadata::Value::I32(v) => v,
                windows_metadata::Value::U32(v) => v as i32,
                _ => 0,
            };
            members.push(EnumMember { name, value });
        }
    }
    TypeMeta::Enum {
        namespace: def.namespace().to_string(),
        name: def.name().to_string(),
        underlying: Box::new(TypeMeta::I32),
        members,
    }
}

fn map_winmd_type(ty: &windows_metadata::Type, index: &reader::Index) -> TypeMeta {
    map_winmd_type_with_generics(ty, index, &[])
}

fn map_winmd_type_with_generics(
    ty: &windows_metadata::Type,
    index: &reader::Index,
    generic_args: &[TypeMeta],
) -> TypeMeta {
    use windows_metadata::Type;
    match ty {
        Type::Void => TypeMeta::Object,
        Type::Bool => TypeMeta::Bool,
        Type::I8 => TypeMeta::I8,
        Type::U8 => TypeMeta::U8,
        Type::I16 => TypeMeta::I16,
        Type::U16 => TypeMeta::U16,
        Type::I32 => TypeMeta::I32,
        Type::U32 => TypeMeta::U32,
        Type::I64 => TypeMeta::I64,
        Type::U64 => TypeMeta::U64,
        Type::F32 => TypeMeta::F32,
        Type::F64 => TypeMeta::F64,
        Type::Char => TypeMeta::Char16,
        Type::String => TypeMeta::String,
        Type::Object => TypeMeta::Object,

        Type::Generic(n) => {
            if (*n as usize) < generic_args.len() {
                generic_args[*n as usize].clone()
            } else {
                TypeMeta::Object
            }
        }

        Type::Name(tn) => resolve_named_type(&tn.namespace, &tn.name, &tn.generics, index),

        Type::Array(inner) | Type::ArrayRef(inner) => {
            TypeMeta::Array(Box::new(map_winmd_type_with_generics(inner, index, generic_args)))
        }

        _ => TypeMeta::Object,
    }
}

fn resolve_named_type(
    namespace: &str,
    name: &str,
    generics: &[windows_metadata::Type],
    index: &reader::Index,
) -> TypeMeta {
    // System.Guid — not in Windows.winmd, handle as primitive
    if namespace == "System" && name == "Guid" {
        return TypeMeta::Guid;
    }

    // Well-known async types
    if namespace == "Windows.Foundation" {
        match name {
            "IAsyncAction" => return TypeMeta::AsyncAction,
            "IAsyncOperation`1" if generics.len() == 1 => {
                return TypeMeta::AsyncOperation(Box::new(map_winmd_type(&generics[0], index)));
            }
            "IAsyncActionWithProgress`1" if generics.len() == 1 => {
                return TypeMeta::AsyncActionWithProgress(Box::new(map_winmd_type(
                    &generics[0],
                    index,
                )));
            }
            "IAsyncOperationWithProgress`2" if generics.len() == 2 => {
                return TypeMeta::AsyncOperationWithProgress(
                    Box::new(map_winmd_type(&generics[0], index)),
                    Box::new(map_winmd_type(&generics[1], index)),
                );
            }
            _ => {}
        }
    }

    // Parameterized interface (generics non-empty)
    if !generics.is_empty() {
        // Try GuidAttribute first, fall back to well-known PIIDs
        let mut piid = match index.get(namespace, name).next() {
            Some(d) => extract_iid(&d),
            None => String::new(),
        };
        if piid.is_empty() {
            piid = well_known_piid(namespace, name).unwrap_or_default().to_string();
        }
        let args = generics.iter().map(|g| map_winmd_type(g, index)).collect();
        return TypeMeta::Parameterized {
            namespace: namespace.to_string(),
            name: name.to_string(),
            piid,
            args,
        };
    }

    let def = match index.get(namespace, name).next() {
        Some(d) => d,
        None => {
            return TypeMeta::Interface {
                namespace: namespace.to_string(),
                name: name.to_string(),
                iid: String::new(),
            };
        }
    };

    if let Some(extends) = def.extends() {
        if extends.namespace() == "System" && extends.name() == "ValueType" {
            let fields = def
                .fields()
                .map(|f| crate::types::FieldMeta {
                    name: f.name().to_string(),
                    typ: map_winmd_type(&f.ty(), index),
                })
                .collect();
            return TypeMeta::Struct {
                namespace: namespace.to_string(),
                name: name.to_string(),
                fields,
            };
        }
        if extends.namespace() == "System" && extends.name() == "Enum" {
            return parse_enum_def(&def);
        }
        if extends.namespace() == "System" && extends.name() == "Object" {
            let default_iid = find_default_interface_iid(&def, index);
            return TypeMeta::RuntimeClass {
                namespace: namespace.to_string(),
                name: name.to_string(),
                default_iid,
            };
        }
    }

    let iid = extract_iid(&def);
    TypeMeta::Interface {
        namespace: namespace.to_string(),
        name: name.to_string(),
        iid,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const WINDOWS_WINMD: &str =
        r"C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.26100.0\Windows.winmd";

    #[test]
    fn test_parse_uri_class() {
        let class = parse_class(WINDOWS_WINMD, "Windows.Foundation", "Uri").unwrap();
        assert_eq!(class.name, "Uri");
        assert_eq!(class.namespace, "Windows.Foundation");
        assert!(class.default_interface.is_some());
        assert!(!class.factory_interfaces.is_empty());
    }

    #[test]
    fn test_uri_vtable_indices() {
        let class = parse_class(WINDOWS_WINMD, "Windows.Foundation", "Uri").unwrap();
        let default_iface = class.default_interface.as_ref().unwrap();
        let scheme = default_iface.methods.iter().find(|m| m.name == "get_SchemeName").unwrap();
        assert!(scheme.is_property_getter);
        let port = default_iface.methods.iter().find(|m| m.name == "get_Port").unwrap();
        assert_eq!(port.return_type, Some(TypeMeta::I32));
    }

    #[test]
    fn test_uri_iid_not_empty() {
        let class = parse_class(WINDOWS_WINMD, "Windows.Foundation", "Uri").unwrap();
        let default_iface = class.default_interface.as_ref().unwrap();
        assert!(!default_iface.iid.is_empty());
    }

    #[test]
    fn test_httpclient_has_default_constructor() {
        let class = parse_class(WINDOWS_WINMD, "Windows.Web.Http", "HttpClient").unwrap();
        assert!(class.has_default_constructor, "HttpClient should have a default constructor");
    }

    #[test]
    fn test_httpclient_overloads_disambiguated() {
        let class = parse_class(WINDOWS_WINMD, "Windows.Web.Http", "HttpClient").unwrap();
        let default_iface = class.default_interface.as_ref().unwrap();
        let names: Vec<&str> = default_iface.methods.iter().map(|m| m.name.as_str()).collect();
        // Should have GetWithOptionAsync, not duplicate GetAsync
        assert!(names.contains(&"GetWithOptionAsync"));
        assert!(names.contains(&"SendRequestWithOptionAsync"));
    }
}

#[cfg(test)]
mod iface_tests {
    use windows_metadata::{HasAttributes, reader};
    const WINDOWS_WINMD: &str =
        r"C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.26100.0\Windows.winmd";

    #[test]
    fn debug_ihttpcontent() {
        let index = reader::Index::read(WINDOWS_WINMD).unwrap();
        let def = index.expect("Windows.Web.Http", "IHttpContent");
        // skip category
        println!("IID: {}", super::extract_iid(&def));
        for (i, m) in def.methods().enumerate() {
            println!("  [{}] {}", i, m.name());
        }
    }
}

