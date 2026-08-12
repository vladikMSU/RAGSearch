# Microsoft MAPIStubLibrary headers

This directory vendors the Extended MAPI headers required to compile
`native-mapi-probe` from a clean checkout.

- Upstream: <https://github.com/microsoft/MAPIStubLibrary>
- Commit: `a9505d73351554078431fc950a0bc34ada6fe39b`
- Upstream commit date: `2026-08-07T08:56:57-04:00`
- Retrieved: 2026-08-12
- License: MIT; see [`LICENSE`](LICENSE)
- Vendored content: the minimal transitive header set required by
  `native-mapi-probe` and the upstream `LICENSE`

Vendored headers:

- `MAPIX.h`
- `MAPIUtil.h`
- `MAPITags.h`
- `MAPIDefS.h`
- `MAPICode.h`
- `MAPIGuid.h`

The upstream library implementation is not vendored or linked. RAGSearch still
links the Windows SDK import library `MAPI32.lib`; at runtime the Windows MAPI
stub dispatches to the MAPI subsystem installed by classic Outlook.

To refresh this dependency intentionally:

1. Select and record an upstream commit.
2. Replace the six listed headers and `LICENSE` with files from that exact commit.
3. Update the commit and retrieval date above.
4. Rebuild `native-mapi-probe` and run the Python/native adapter tests.
