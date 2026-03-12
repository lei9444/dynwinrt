/**
 * End-to-end test using 100% winrt-meta generated bindings + dynwinrt-js.
 * Tests: create picker, set properties, use IVector_String.append(), open dialog.
 */
import { initWinappsdk, DynWinRtValue } from 'dynwinrt-js'
import { FileOpenPicker } from './generated/FileOpenPicker'
import { PickerViewMode } from './generated/PickerViewMode'

async function main() {
    initWinappsdk(1, 8)

    // Create picker
    const picker = FileOpenPicker.createInstance(DynWinRtValue.i64(0))
    console.log('FileOpenPicker created')

    // Set properties
    picker.viewMode = PickerViewMode.Thumbnail
    console.log('ViewMode:', picker.viewMode, '(expected:', PickerViewMode.Thumbnail, ')')
    console.assert(picker.viewMode === PickerViewMode.Thumbnail, 'ViewMode mismatch')

    picker.commitButtonText = 'Select File'
    console.log('CommitButtonText:', picker.commitButtonText)
    console.assert(picker.commitButtonText === 'Select File', 'CommitButtonText mismatch')

    // Use generated IVector_String — fully typed, no DynWinRtValue wrapping!
    const filter = picker.fileTypeFilter
    filter.append('.png')
    filter.append('.jpg')
    filter.append('.txt')
    console.log('FileTypeFilter size:', filter.size)
    console.assert(filter.size === 3, 'Expected 3 filters')

    console.log('Filter[0]:', filter.getAt(0))
    console.log('Filter[1]:', filter.getAt(1))
    console.log('Filter[2]:', filter.getAt(2))

    // Open file picker dialog
    console.log('Opening file picker dialog...')
    const result = await picker.pickSingleFileAsync()
    if (result && result._obj) {
        console.log('Selected file path:', result.path)
    } else {
        console.log('User cancelled the picker')
    }

    console.log('ALL PASS')
}

main().catch(e => console.error('Error:', e))
