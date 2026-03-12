/**
 * Basic test using Windows SDK only — no WinAppSDK needed.
 * Run: ./winrt-meta.exe generate --namespace "Windows.Foundation" --class-name "Uri" --output ./generated-uri
 * Then: npm install && npx tsx test_uri.ts
 */
import { roInitialize } from 'dynwinrt-js'
import { Uri } from './generated-uri/Uri'

roInitialize(1)

const uri = Uri.createUri('https://www.example.com:8080/path?q=test#frag')
console.log('AbsoluteUri:', uri.absoluteUri)
console.log('Host:', uri.host)
console.log('Port:', uri.port)
console.log('Path:', uri.path)
console.log('Query:', uri.query)
console.log('SchemeName:', uri.schemeName)

console.assert(uri.host === 'www.example.com', 'Host mismatch')
console.assert(uri.port === 8080, 'Port mismatch')
console.assert(uri.path === '/path', 'Path mismatch')
console.assert(uri.schemeName === 'https', 'Scheme mismatch')

const combined = Uri.createUri('https://example.com/api/').combineUri('v1/users')
console.log('Combined:', combined.absoluteUri)
console.assert(combined.absoluteUri === 'https://example.com/api/v1/users', 'CombineUri mismatch')

console.log('ALL PASS')
