# Dynamic Projection For JS/Python

## Objective

To expose WinRT Windows APIs to dynamic languages like JavaScript and Python with a seamless developer experience, avoiding the need for native compilation (C++ compilers) and strict version coupling between the projection and the Windows App SDK (WASDK) components.

## Motivation & Challenges

- **Versioning Hell**: Current static projections (e.g., existing PyWinRT) require generating and compiling native code for specific versions of WinRT components. This creates a matrix of compatibility issues and requires releasing new projection packages for every WASDK release.
- **Developer Experience (DX)**: Node.js/Electron and Python developers expect "pip install" or "npm install" experiences without needing MSVC tools or C++ compilers in their build chain.
- Static vs. Dynamic Nature:
  - **Static Languages (C++, Rust, C#)**: Have full type information at compile time. Static projections can "tailor" the bindings to exactly what is used.
  - **Dynamic Languages (JS, Python)**: Rely on runtime (lazy) evaluation. Objects are often type-erased and passed around. A static projection leaves "holes" when a function receives an unknown WinRT object at runtime that wasn't statically projected.
  - TODO: details on why static projection is not ideal for dynamic languages, if only concern is trimming, maybe some tree-shaking features can help?

## Architecture

The proposed architecture moves away from static native bindings towards a dynamic approach. It consists of two main separable components:

### The Runtime Library

A minimal runtime library native to the target ecosystem (e.g., a `.pyd` or Node addon) that enables dynamic calls to arbitrary WinRT APIs.

- **FFI & ABI Handling**:
  A minimum dynamic FFI layer that can:
  - Call arbitrary WinRT methods via function pointer and proper parameter type information, typically using dyncall or libffi.
  - Managing `out` parameters (stack allocation).
  - Direct Value type pass by value support, computing struct sizes and alignments at runtime (since `sizeof` isn't available).
- **Platform Primitives**: Wraps fundamental OS API (strings, `RoInitialize`, `RoGetActivationFactory`, `QueryInterface`).
- **Language Adaptation**:
  - Maps WinRT `HSTRING`s to language-native strings.
  - Maps WinRT `HRESULT`s to language exceptions
  - `IAsyncAction` to language-specific Promises/Awaitables.
    **WinAppSDK Bootstrap**: Include proper bootstrap dll and native interop, to enable WinAppSDK usage for unpackaged apps.

### Metadata Parser & Projection Generator

This component bridges the gap between the raw WinMD metadata and the runtime projection. There are two modes of operation:

- **Mode A: Fully Lazy (Runtime)**
  - The runtime parses `.winmd` files on the fly as APIs are accessed.
  - **Pros**: Battle tested in old JS projection. Simpler distribution (no generation step).
  - **Cons**: runtime parsing overhead (though potentially negligible compared to marshalling). No IDE intellisense.
- **Mode B: Design-Time Generation (Pre-processed)**
  - A CLI tool parses `.winmd` files and generates **non-native** code (pure `.js` or `.py` files) that describes the interface shapes and method signatures for the runtime.
  - Can generates IDE helpers (e.g., TypeScript `.d.ts` files, and Python `.pyi` stub files).
  - **Pros**: Better DX (IntelliSense/Autocomplete), potentially faster startup (no WinMD parsing).
  - **Cons**: Requires a generation step (but _not_ a native compilation step).

## Workflow

1.  **Runtime**: The runtime is shipped as a generic library for the language (e.g., `pip install lazy-winrt`).
2.  **Projection**:
    - In theory, the runtime can directly support runtime interface specification.
    - Developers can use the libraries directly (lazy loading namespaces) with winmd distributed, and the runtime will parse winmd and generated the necessary interface shapes on the fly.
    - OR developers run a tool (e.g. `npx lazy-winrt-gen`) to generate bindings/types for the specific WinMDs they use.

## Performance Considerations

- **Overhead**: The primary cost is expected to be in WinMD parsing (if lazy) and the dynamic WinRT methods invocation overhead (dynamic dispatch).
- **Comparison**: This method calling overhead is likely comparable to the existing marshalling costs of crossing the JS/Python boundary.
- **Optimization**: A hybrid approach (Design-Time generation of interface shapes) eliminates the WinMD parsing cost at runtime, leaving only the FFI overhead.

## Implementation Details

The runtime provides a minimum representation of WinRT types/values and conversions of WinRT values to language-native values.
The runtime provide a minimum interface specification language to the target language(JS, Python), allowing the user to define WinRT interfaces/classes at runtime, with following sample syntax

```ts
import { WinRT } from "lazy-winrt";

const UriInterface = WinRT.Interface({
  namespace: "Windows.Foundation",
  name: "IUriRuntimeClass",
  guid: "<...guid of the interface...>",
  methods: [
    // get AbsoluteUri(): string
    WinRT.Method([WinRT.Out(WinRT.HSTRING)]), // Method implicitly assume first argument is ComPtr, and result is HRESULT
    // get Domain(): string
    WinRT.Method([WinRT.Out(WinRT.HSTRING)]),
    // ... other methods
  ],
});
const UriFactoryInterface = WinRT.Interface({
  namespace: "Windows.Foundation",
  name: "IUriRuntimeClassFactory",
  guid: "<...guid of the interface...>",
  methods: [
    // CreateUri(string uri): IUriRuntimeClass
    WinRT.Method([WinRT.HSTRING, WinRT.Out(UriInterface)]),
  ],
});

class Uri {
  // statically cached factory object, but in target dynamic language instead of static native code
  static Factory = WinRT.as(UriFactoryInterface, WinRT.getActivationFactory("Windows.Foundation.Uri"));

  constructor(uriString: string) {
    this._instance = WinRT.activateInstance("Windows.Foundation.Uri", [
      uriString,
    ]);
  }

  get absoluteUri(): string {
    return WinRT.callMethod(UriInterface, 6, [this._instance]); // 6 is the vtable index of get_AbsoluteUri
  }
}
```

winmd parser can generate interface specification / developer friendly projection classes at runtime.

### Stub Method Optimizations

Instead of calling all methods using dynamical FFI calls, we can put some commonly used method signatures in the runtime library, so that for those methods, they are almost as fast as static projection.
For those stub methods, only actual ABI signature matters, thus a large kinds of WinRT methods can map to same stub method, e.g. all out parameters are just treated as same pointer type, lots of getter like/factory like methods would just map to stub signatures like following:
```
// same stub for all object.get_X -> Com/HSTRING like references types
HRESULT Method_Out_Pointer(void* funPtr, ComPtr self, void* outValue) {
    var f = // cast funPtr to proper function pointer type
    return f(self, outValue);
}
```

### Challenges

* Properly handling signature (guids) casting, especially for generics, also need special handling of their JS/Python representation, e.g. `IVector` may need to map to a function instead of a simple interface instance.
* Handling of async, although 

## References

- **Legacy JS Projection**: Old JS apps used a dynamic projection where the runtime read WinMDs. Perf was acceptable.
- **PyWinRT**: Earlier experiments with static C++/WinRT-based projections highlighted the versioning and distribution difficulties.
- **"Lazy-WinRT" Prototype**: Shows that parsing WinMD and invoking methods dynamically is feasible and potentially performant enough.
- **dynwinrt** a Rust based implementation inspired by Lazy-WinRT to leverage NAPI-RS/PyO3 for easier integration with JavaScript and Python.
