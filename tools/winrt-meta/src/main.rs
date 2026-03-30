use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::Path;

use clap::{Parser, Subcommand};

use winrt_meta::codegen::typescript;
use winrt_meta::meta;
use winrt_meta::types::TypeMeta;

#[derive(Parser)]
#[command(name = "winrt-meta")]
#[command(about = "Generate typed language bindings from WinRT metadata (.winmd) files")]
#[command(long_about = "winrt-meta reads .winmd metadata and generates typed TypeScript bindings\n\
    that use dynwinrt-js at runtime to call Windows Runtime APIs dynamically.\n\n\
    It auto-detects Windows SDK metadata and discovers sibling .winmd files\n\
    in the same directory, so you typically only need to point at one file.")]
#[command(after_help = "\x1b[1mExamples:\x1b[0m\n\
    # Generate all namespaces from a WinAppSDK metadata folder\n\
    winrt-meta generate --folder C:\\Users\\you\\.winapp\\packages\\Microsoft.WindowsAppSDK.AI.1.8.39\\metadata\n\n\
    # Generate a single namespace (siblings auto-discovered)\n\
    winrt-meta generate --winmd path\\to\\Microsoft.Windows.AI.Imaging.winmd --namespace Microsoft.Windows.AI.Imaging\n\n\
    # Generate a single class\n\
    winrt-meta generate --namespace Windows.Foundation --class Uri\n\n\
    # Custom output directory\n\
    winrt-meta generate --folder path\\to\\metadata --output ./src/generated")]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// Generate TypeScript bindings from .winmd files
    #[command(long_about = "Parse .winmd metadata and generate typed TypeScript files.\n\n\
        The tool automatically:\n\
        - Detects Windows.winmd from the Windows SDK install path\n\
        - Discovers sibling .winmd files in the same directory as --winmd\n\
        - Resolves transitive type dependencies across namespaces\n\
        - Filters out Windows.* system namespaces when --namespace is omitted")]
    Generate {
        /// Path(s) to .winmd metadata files, separated by ';'.
        /// Sibling .winmd files in the same directory are auto-discovered.
        /// If omitted, auto-detects Windows.winmd from Windows SDK.
        #[arg(long, value_name = "PATH")]
        winmd: Option<String>,

        /// Directory containing .winmd files.
        /// All .winmd files in this directory will be loaded.
        /// When --namespace is omitted, generates all non-Windows namespaces.
        #[arg(long, value_name = "DIR")]
        folder: Option<String>,

        /// Generate only this namespace (e.g. "Microsoft.Windows.AI.Imaging").
        /// If omitted, generates all non-Windows namespaces found in the winmd files.
        #[arg(long, value_name = "NS")]
        namespace: Option<String>,

        /// Generate bindings for a single class (requires --namespace)
        #[arg(long, name = "class", value_name = "NAME")]
        class_name: Option<String>,

        /// Target language
        #[arg(long, default_value = "ts", value_parser = ["ts"])]
        lang: String,

        /// Output directory for generated files
        #[arg(long, default_value = "./generated", value_name = "DIR")]
        output: String,
    },
}

fn main() {
    let cli = Cli::parse();

    match cli.command {
        Commands::Generate {
            winmd,
            folder,
            namespace,
            class_name,
            lang: _,
            output,
        } => {
            // Collect winmd paths from --folder and/or --winmd
            let mut winmd_parts: Vec<String> = Vec::new();

            if let Some(ref dir) = folder {
                let dir_path = Path::new(dir);
                if !dir_path.is_dir() {
                    eprintln!("--folder path is not a directory: {}", dir);
                    std::process::exit(1);
                }
                if let Ok(entries) = fs::read_dir(dir_path) {
                    for entry in entries.flatten() {
                        let path = entry.path();
                        if path.extension().map_or(false, |ext| ext.eq_ignore_ascii_case("winmd")) {
                            eprintln!("Loading winmd from folder: {}", path.display());
                            winmd_parts.push(path.to_string_lossy().to_string());
                        }
                    }
                }
                if winmd_parts.is_empty() {
                    eprintln!("No .winmd files found in folder: {}", dir);
                    std::process::exit(1);
                }
            }

            if let Some(ref w) = winmd {
                winmd_parts.extend(w.split(';').filter(|s| !s.is_empty()).map(String::from));
            }

            // Auto-detect Windows SDK if not already included
            let has_windows_winmd = winmd_parts.iter().any(|p| p.contains("Windows.winmd"));
            if !has_windows_winmd {
                if let Some(sdk_winmd) = find_windows_sdk_winmd() {
                    eprintln!("Auto-detected Windows SDK: {}", sdk_winmd);
                    winmd_parts.push(sdk_winmd);
                } else if folder.is_none() && winmd.is_none() {
                    eprintln!("Could not auto-detect Windows.winmd. Please provide --winmd or --folder.");
                    std::process::exit(1);
                }
            }

            let winmd_joined = winmd_parts.join(";");

            // Auto-discover sibling .winmd files in the same directories
            let winmd = meta::expand_winmd_paths(&winmd_joined);

            let output_dir = Path::new(&output);
            fs::create_dir_all(output_dir).expect("Failed to create output directory");

            if let Some(ref cls) = class_name {
                // Single class mode
                let ns = namespace
                    .as_deref()
                    .expect("--namespace is required when --class is specified");
                let classes = match meta::parse_class(&winmd, ns, cls) {
                    Some(c) => vec![c],
                    None => {
                        eprintln!("Class {}.{} not found in {}", ns, cls, winmd);
                        std::process::exit(1);
                    }
                };
                let _ = generate_for_types(&winmd, output_dir, classes.clone(), Vec::new(), Vec::new());

                // Append to existing index file if present
                let index_path = output_dir.join("index.ts");
                if index_path.exists() {
                    let deps = meta::resolve_dependencies(&winmd, &classes, &[], &[]);
                    let all_classes = [classes.as_slice(), deps.classes.as_slice()].concat();
                    let existing = fs::read_to_string(&index_path).expect("Failed to read index file");
                    let updated = typescript::append_to_index(&existing, &all_classes, &deps.interfaces, &deps.enums);
                    fs::write(&index_path, &updated).expect("Failed to update index file");
                    println!("Updated {}", index_path.display());
                }
            } else {
                // Determine which namespaces to generate
                let namespaces = match namespace {
                    Some(ref ns) => vec![ns.clone()],
                    None => {
                        let all_ns = meta::list_namespaces(&winmd);
                        let filtered: Vec<String> = all_ns
                            .into_iter()
                            .filter(|ns| !ns.starts_with("Windows."))
                            .collect();
                        if filtered.is_empty() {
                            eprintln!("No non-Windows namespaces found. Use --namespace to specify one.");
                            std::process::exit(1);
                        }
                        eprintln!("Discovered {} namespace(s) to generate:", filtered.len());
                        for ns in &filtered {
                            eprintln!("  {}", ns);
                        }
                        filtered
                    }
                };

                let mut total_classes = 0usize;
                let mut total_interfaces = 0usize;
                let mut total_enums = 0usize;

                for ns in &namespaces {
                    let classes = meta::parse_namespace(&winmd, ns);
                    let interfaces = meta::parse_interfaces(&winmd, ns);
                    let enums = meta::parse_enums(&winmd, ns);

                    let (nc, ni, ne) = generate_for_types(
                        &winmd, output_dir, classes, interfaces, enums,
                    );
                    total_classes += nc;
                    total_interfaces += ni;
                    total_enums += ne;
                }

                // Generate index file combining everything
                if namespaces.len() >= 1 && (total_classes + total_interfaces + total_enums) > 1 {
                    let mut all_classes = Vec::new();
                    let mut all_interfaces = Vec::new();
                    let mut all_enums = Vec::new();
                    for ns in &namespaces {
                        all_classes.extend(meta::parse_namespace(&winmd, ns));
                        all_interfaces.extend(meta::parse_interfaces(&winmd, ns));
                        all_enums.extend(meta::parse_enums(&winmd, ns));
                    }
                    let deps = meta::resolve_dependencies(&winmd, &all_classes, &all_interfaces, &all_enums);
                    all_classes.extend(deps.classes);
                    all_interfaces.extend(deps.interfaces);
                    all_enums.extend(deps.enums);

                    let index_code = typescript::generate_index(&all_classes, &all_interfaces, &all_enums);
                    let index_path = output_dir.join("index.ts");
                    fs::write(&index_path, &index_code).expect("Failed to write index file");
                    println!("Generated {}", index_path.display());
                }

                println!(
                    "Done. {} class(es) + {} interface(s) + {} enum(s) generated in {}",
                    total_classes, total_interfaces, total_enums, output_dir.display()
                );
            }
        }
    }
}

/// Generate .ts files for a set of types plus their transitive dependencies.
fn generate_for_types(
    winmd: &str,
    output_dir: &Path,
    classes: Vec<meta::ClassMeta>,
    interfaces: Vec<meta::InterfaceMeta>,
    enums: Vec<TypeMeta>,
) -> (usize, usize, usize) {
    let deps = meta::resolve_dependencies(winmd, &classes, &interfaces, &enums);
    let mut all_classes = classes;
    let mut all_interfaces = interfaces;
    let mut all_enums = enums;
    all_classes.extend(deps.classes);
    all_interfaces.extend(deps.interfaces);
    all_enums.extend(deps.enums);

    let mut known_types: HashSet<String> = HashSet::new();
    for c in &all_classes { known_types.insert(c.name.clone()); }
    for i in &all_interfaces { known_types.insert(i.name.clone()); }
    for e in &all_enums {
        if let TypeMeta::Enum { name, .. } = e { known_types.insert(name.clone()); }
    }

    let delegate_type_names: HashSet<String> = all_interfaces.iter()
        .filter(|i| i.methods.iter().any(|m| m.name == ".ctor") && i.methods.iter().any(|m| m.name == "Invoke"))
        .map(|i| i.name.clone())
        .collect();

    let mut req_iface_count: HashMap<String, (&meta::InterfaceMeta, usize)> = HashMap::new();
    for class in &all_classes {
        for ri in &class.required_interfaces {
            if ri.iid.is_empty() { continue; }
            req_iface_count.entry(ri.iid.clone())
                .and_modify(|(_, c)| *c += 1)
                .or_insert((ri, 1));
        }
    }
    let shared_iids: HashSet<String> = req_iface_count.iter()
        .filter(|(_, (_, count))| *count >= 2)
        .map(|(iid, _)| iid.clone())
        .collect();

    let shared_interfaces: Vec<meta::InterfaceMeta> = req_iface_count.iter()
        .filter(|(_, (_, count))| *count >= 2)
        .map(|(_, (iface, _))| (*iface).clone())
        .collect();
    for iface in &shared_interfaces {
        known_types.insert(iface.name.clone());
    }

    // Generate shared interfaces
    for iface in &shared_interfaces {
        let code = typescript::generate_interface(iface, &known_types, &delegate_type_names);
        let filepath = output_dir.join(format!("{}.ts", iface.name));
        fs::write(&filepath, &code).expect("Failed to write shared interface file");
        println!("Generated shared {}", filepath.display());
    }

    // Generate interfaces
    for iface in &all_interfaces {
        let code = typescript::generate_interface(iface, &known_types, &delegate_type_names);
        let filepath = output_dir.join(format!("{}.ts", iface.name));
        fs::write(&filepath, &code).expect("Failed to write generated file");
        println!("Generated {}", filepath.display());
    }

    // Generate enums
    for en in &all_enums {
        if let TypeMeta::Enum { name, .. } = en {
            if let Some(code) = typescript::generate_enum(en) {
                let filepath = output_dir.join(format!("{}.ts", name));
                fs::write(&filepath, &code).expect("Failed to write generated file");
                println!("Generated {}", filepath.display());
            }
        }
    }

    // Generate classes
    for class in &all_classes {
        let code = typescript::generate_class(class, &known_types, &delegate_type_names, &shared_iids);
        let filepath = output_dir.join(format!("{}.ts", class.name));
        fs::write(&filepath, &code).expect("Failed to write generated file");
        println!("Generated {}", filepath.display());
    }

    (all_classes.len(), all_interfaces.len(), all_enums.len())
}

/// Find Windows SDK Windows.winmd by scanning the standard install location.
fn find_windows_sdk_winmd() -> Option<String> {
    let base = Path::new(r"C:\Program Files (x86)\Windows Kits\10\UnionMetadata");
    if !base.exists() {
        return None;
    }
    let mut versions: Vec<_> = fs::read_dir(base)
        .ok()?
        .filter_map(|e| e.ok())
        .filter(|e| e.path().is_dir())
        .map(|e| e.file_name().to_string_lossy().to_string())
        .filter(|name| name.starts_with("10."))
        .collect();
    versions.sort();
    for version in versions.iter().rev() {
        let winmd_path = base.join(version).join("Windows.winmd");
        if winmd_path.exists() {
            return Some(winmd_path.to_string_lossy().to_string());
        }
    }
    None
}
