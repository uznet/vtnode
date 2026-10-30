# vtnode_httpa_params.ini configuration guide

[한국어](./vtnode_httpa_params.ko.md) | **English**

Reviewed: 2026-10-22. Source: [vtnode_httpa_params.ini](./vtnode_httpa_params.ini) in this folder.

## Evidence and scope

The values below are **current values read directly from the source file**, not Runtime defaults for omitted settings. Descriptions follow the INI comments and key names. License instructions were cross-checked against the local [free-license guide](./FREE_LICENSE.md). The native INI parser and execution engine implementation were not verified, and no runtime behavior tests were performed. Uncertain details are marked explicitly.

## Current configuration

- HTTP port: `12345`. HTTPS disabled through `ssl.port=0`.
- WebSocket enabled; message limit 1MB and frame limit 64KB.
- Host policy: `example.com`, with aliases `www.example.com` and `api.example.com`; arbitrary Host acceptance disabled.
- Debug output enabled, stdout console output disabled, license use disabled.
- `maxConnections=16` does not establish entitlement to 16 simultaneous connections. The local license guide describes unlicensed operation as one channel; the exact relationship between this setting and licensed channel limits is unverified.

## Syntax and units

- `[section]` introduces a section; `key=value` specifies a value. Preserve dotted section/key names such as `httpserver.timeout` and `ssl.port`.
- This file uses `;` for comment lines and trailing comments. `;names= chat,echo` is not an active setting.
- The source defines `1KB=1024` bytes and `1MB=1024×1024` bytes: 4KB=4,096; 8KB=8,192; 64KB=65,536; 1MB=1,048,576; 10MB=10,485,760 bytes.
- Comments identify `10s`, `30s`, `60s`, and `600s` as seconds. Parser rules for other units, case, whitespace, and quoting have not been verified.
- Zero has a key-specific meaning. Treat it as unlimited/disabled only where explicitly documented below.
- Idle timeouts limit periods without progress, as distinct from overall elapsed-time limits. Internal timer reset conditions remain unverified.

## Settings

### [http] — HTTP request limits

| Key | Current value | Description |
| --- | --- | --- |
| `maxRequestLineBytes` | `4KB` | Maximum request-line size. The comment recommends including CRLF. |
| `maxHeaderLineBytes` | `4KB` | Maximum size of one header line, including CRLF. |
| `maxHeaderCount` | `100` | Maximum number of allowed headers. |
| `maxHeaderTotalBytes` | `8KB` | Maximum header-section size, excluding the request line and including CRLFCRLF. |
| `maxBodyBytes` | `10MB` | Maximum request-body size. |
| `maxTransactionsPerConnection` | `10` | Maximum transactions per connection; 0 means unlimited. |


### [httpserver.timeout] — HTTP server timeouts

| Key | Current value | Description |
| --- | --- | --- |
| `requestFirstByteTimeoutSec` | `10s` | Maximum wait for the first byte of the first request line. |
| `requestHeaderTimeoutSec` | `30s` | Time from receiving the first byte to completing all request headers. |
| `requestBodyIdleTimeoutSec` | `60s` | Idle limit while receiving the request body. |
| `transactionTimeoutSec` | `60s` | Limit from request completion to completion of response headers and body. |
| `keepAliveTimeoutSec` | `60s` | Wait for the next request on the same connection after completing a response. |


### [httpclient.timeout] — HTTP client timeouts

| Key | Current value | Description |
| --- | --- | --- |
| `requestSendIdleTimeoutSec` | `30s` | Idle limit while writing the request. |
| `responseHeaderTimeoutSec` | `30s` | Comment describes the wait from request completion to response headers/first byte (TTFB). The exact stopping condition needs implementation verification. |
| `responseBodyIdleTimeoutSec` | `60s` | Idle limit while receiving the response body. |
| `overallTimeoutSec` | `0` | Overall round-trip limit. The current 0 disables this limit; individual timeouts still apply. |


### [websocket] — WebSocket enablement and size limits

| Key | Current value | Description |
| --- | --- | --- |
| `enabled` | `true` | WebSocket enabled; false disables it. |
| `maxMessageBytes` | `1MB` | Maximum complete message size after reassembling fragmented frames. |
| `maxFrameBytes` | `64KB` | Maximum single-frame size. |
| `maxMissingPongCount` | `1` | Allowed consecutive missing pong count. The exact disconnect threshold comparison is unverified. |


### [websocket.timeout] — WebSocket timeouts

| Key | Current value | Description |
| --- | --- | --- |
| `closeTimeoutSec` | `10s` | Wait for the peer close frame after sending a close frame. |
| `idleTimeoutSec` | `600s` | Idle limit with no frames exchanged (10 minutes). |
| `pongTimeoutSec` | `10s` | Wait for pong after ping. This file does not specify a ping interval. |


### [websocket.subprotocol] — WebSocket subprotocol negotiation

| Key | Current value | Description |
| --- | --- | --- |
| `policy` | `optional` | The comment lists optional / required. The current setting makes subprotocol negotiation optional. |
| `serverPreference` | `true` | Prefer server selection. false turns off server preference; the detailed selection order is unverified. |
| `onUnsupported` | `ignore` | Handling unsupported subprotocols: the comment lists ignore and reject. |

Commented example: `;names= chat,echo`. No supported-name list is explicitly active. Built-in behavior when the list is omitted, precedence between `required` and `ignore`, and negotiation-failure responses need implementation verification.


### [net] — Network and TLS

| Key | Current value | Description |
| --- | --- | --- |
| `maxConnections` | `16` | Named as a maximum connection setting. The comment is empty; its scope and precedence relative to licensing are unverified. |
| `port` | `12345` | Plain HTTP port; enabled when greater than 0 according to the comment. |
| `ssl.port` | `0` | HTTPS port. Currently disabled; a value greater than 0 enables it. |
| `ssl.certFile` | `server.crt` | Server certificate file path. The base directory for relative paths is unverified. |
| `ssl.keyFile` | `server.key` | Server private-key file path. The base directory for relative paths is unverified. |


### [log] — Log output

| Key | Current value | Description |
| --- | --- | --- |
| `debugOut` | `true` | Debug output enabled. Its destination and relationship to file logging are not specified here. |
| `consoleOut` | `false` | Console logging to standard output (stdout) disabled. |


### [host.policy] — Server Host policy

| Key | Current value | Description |
| --- | --- | --- |
| `primaryHost` | `example.com` | Primary host name; currently an example domain. |
| `allowAnyHost` | `false` | Arbitrary Host acceptance disabled. true accepts regardless of the Host header (commented for proxy/testing use). |
| `aliasHosts` | `www.example.com , api.example.com` | Additional host names; currently two comma-separated entries. |


### [license] — License

| Key | Current value | Description |
| --- | --- | --- |
| `enabled` | `false` | License use disabled. Set true when using an issued license. |
| `id` | `"your-license-id"` | Placeholder for the issued License ID. Supply the ID without the .vtlic extension. |

## Applying and checking settings

1. Identify and back up the INI actually loaded by the application. Do not assume this quick-start copy is automatically used by the running application.
2. Change the required settings, restart the application, and check the results. The local license guide explicitly requires a restart. Hot reload for other settings is unverified, so checking after a restart is recommended.
3. Check both the port and Host policy when accessing HTTP. Requests using `localhost` or `127.0.0.1` may not match the current host list. This manual example targets a local server with an explicit Host header; it was not run during documentation preparation.

```powershell
curl.exe -v -H "Host: example.com" http://127.0.0.1:12345/
```

The response at `/` depends on application routing. Host rejection status codes and case, port, or wildcard matching rules remain unverified.

### Enabling HTTPS

Set `ssl.port` to the intended positive port and provide valid certificate and private-key paths. Paths being present in the current file do not enable HTTPS. This INI does not establish relative-path resolution, certificate formats, TLS versions, or certificate-error handling.

### Using a license

Follow the [free-license guide](./FREE_LICENSE.md) to activate the Meta Code on the target PC, then set `[license]` to `enabled=true` and the issued `id`, and restart the application. According to that guide, Runtime locates `%VTNODE_HOME%\licenses\<License-ID>.vtlic` and validates it at startup. `"your-license-id"` is a placeholder, not an issued ID.

## Troubleshooting reference

| Symptom | First checks |
| --- | --- |
| HTTP connection fails | Running server, `net.port`, port conflicts, firewall |
| localhost access is not accepted | Primary host, aliases, and `allowAnyHost` under `host.policy` |
| Large requests are not processed | Individual request-line, header count, total headers, and body limits |
| Connection or response ends early | Server/client timeouts and transactions-per-connection limit |
| WebSocket negotiation/message processing fails | Enablement, subprotocol policy, frame/message limits, pong/idle limits |
| No console logs | Current `consoleOut=false`; separately verify the actual `debugOut` destination |
| Fewer simultaneous connections than configured | Licensed channels and runtime logs; do not infer capacity from `maxConnections` alone |

HTTP status codes, WebSocket close codes, exact disconnect timing, fallback behavior for missing/invalid values, and accepted ranges require implementation or runtime-log verification. The source INI was not modified in preparing this guide.
