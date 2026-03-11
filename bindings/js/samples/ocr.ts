import {
    DynWinRtValue,
    DynWinRtType,
    DynWinRtMethodSig,
    WinIIds,
    OcrDemo,
    hasPackageIdentity,
} from '../dist/index.js'

// ======================================================================
// Register interfaces (once)
// ======================================================================

// ITextRecognizerStatics: GetReadyState, EnsureReadyAsync, CreateAsync
const iTextRecognizerStatics = DynWinRtType.registerInterface(
    "ITextRecognizerStatics",
    WinIIds.iTextRecognizerStaticsIid(),
)
    .addMethod("GetReadyState", new DynWinRtMethodSig().addOut(DynWinRtType.i32()))
    .addMethod("EnsureReadyAsync", new DynWinRtMethodSig().addOut(DynWinRtType.object()))
    .addMethod("CreateAsync", new DynWinRtMethodSig().addOut(
        DynWinRtType.iAsyncOperation(
            DynWinRtType.runtimeClass(
                'Microsoft.Windows.AI.Imaging.TextRecognizer',
                WinIIds.textRecognizerIid(),
            )
        )))

// IImageBufferStatics: CreateForSoftwareBitmap
const iImageBufferStatics = DynWinRtType.registerInterface(
    "IImageBufferStatics",
    WinIIds.iImageBufferStaticsIid(),
)
    .addMethod("CreateForSoftwareBitmap", new DynWinRtMethodSig()
        .addIn(DynWinRtType.object()).addOut(DynWinRtType.object()))

// ITextRecognizer: RecognizeTextFromImageAsync
const iTextRecognizer = DynWinRtType.registerInterface(
    "ITextRecognizer",
    WinIIds.textRecognizerIid(),
)
    .addMethod("RecognizeTextFromImageAsync", new DynWinRtMethodSig()
        .addIn(DynWinRtType.object())
        .addOut(DynWinRtType.iAsyncOperation(
            DynWinRtType.runtimeClass(
                'Microsoft.Windows.AI.Imaging.RecognizedText',
                WinIIds.recognizedTextIid(),
            )
        )))

// ======================================================================
// Use
// ======================================================================

async function main() {
    // initWinappsdk(1, 8); // no need in packaged app
    console.log(hasPackageIdentity() ? 'Has package identity' : 'No package identity');

    // Create recognizer
    const factory = DynWinRtValue.activationFactory('Microsoft.Windows.AI.Imaging.TextRecognizer')
        .cast(WinIIds.iTextRecognizerStaticsIid())

    const readyState = iTextRecognizerStatics.methodByName("GetReadyState").invoke(factory, [])
    console.log('TextRecognizer ready state:', readyState.toNumber())

    const recognizer = await iTextRecognizerStatics.methodByName("CreateAsync")
        .invoke(factory, []).toPromise()

    // Create image buffer from picked bitmap
    const bitmap = OcrDemo.pickImage()
    const ibFactory = DynWinRtValue.activationFactory('Microsoft.Graphics.Imaging.ImageBuffer')
        .cast(WinIIds.iImageBufferStaticsIid())
    const imageBuffer = iImageBufferStatics.methodByName("CreateForSoftwareBitmap")
        .invoke(ibFactory, [bitmap])

    // Recognize text
    const result = await iTextRecognizer.methodByName("RecognizeTextFromImageAsync")
        .invoke(recognizer, [imageBuffer]).toPromise()

    OcrDemo.printOcrLines(result)
}
main();
