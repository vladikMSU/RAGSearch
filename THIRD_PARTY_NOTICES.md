# Third-party notices

This file covers third-party material used or referenced by RAGSearch. It does not
license the first-party RAGSearch source code. The source repository contains the
scoped commercial-compatibility audit at `docs/LICENSE_COMPATIBILITY.md`.

## Microsoft MAPIStubLibrary

RAGSearch vendors the minimal header subset required by
`connectors/outlook_mapi/native/OutlookMapiReader.vcxproj` from:

- project: <https://github.com/microsoft/MAPIStubLibrary>
- commit: `a9505d73351554078431fc950a0bc34ada6fe39b`
- copyright: Copyright (c) 2018 Microsoft
- license: MIT

The complete license is shipped beside this file as
`MAPIStubLibrary-LICENSE.txt`; its source-repository location is
`third_party/MAPIStubLibrary/LICENSE`. Keep that license with source and binary
distributions containing material derived from these headers. These files remain
MIT-licensed and are not relicensed by any first-party RAGSearch license or EULA.

## Microsoft VSTO redistributable files

The VSTO build copies these files unmodified from Visual Studio:

- `Microsoft.Office.Tools.Common.v4.0.Utilities.dll`
- `Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll`

They appear in Microsoft's
[Visual Studio 2022 redistribution list](https://learn.microsoft.com/visualstudio/releases/2022/redistribution).
Their redistribution is permitted only subject to the applicable Microsoft/Visual
Studio license terms. They are not first-party RAGSearch code and must not be
represented as such.

The release producer must use a properly licensed Visual Studio edition/toolchain
and provide end-user/distributor terms that satisfy Microsoft's redistribution
conditions. The development build's self-signed certificate is not a release
publisher credential.

## External prerequisites not distributed here

Windows, Visual Studio/MSVC, the Windows SDK, .NET Framework, VSTO Runtime, classic
Outlook and the installed MAPI provider are external prerequisites governed by
their own Microsoft terms. The current repository does not redistribute them.

The Python interpreter and standard library are also external prerequisites. If a
future installer bundles Python, it must include the PSF license stack, copyright
notices and incorporated-software notices for the exact Python version.

## Optional neural components not distributed here

`sentence-transformers`, its transitive dependencies, and
`paraphrase-multilingual-MiniLM-L12-v2` are optional, ignored by Git and never
downloaded automatically. The optional direct package version is selected as
`sentence-transformers==5.7.0`, and the package/model identify as Apache-2.0, but
transitive artifacts, hashes and the model revision are not locked. They are
therefore outside the cleared default distribution set; audit the exact lock,
native wheels and model revision before bundling them.

The locally tested `sentence-transformers 5.7.0` distribution contains both an
Apache-2.0 `LICENSE` and an upstream `NOTICE`; both would have to accompany a bundled
copy. The model card identifies Apache-2.0 but is not a substitute for pinning and
packaging the full license text.

Microsoft and Outlook are trademarks of the Microsoft group of companies. RAGSearch
is not affiliated with or endorsed by Microsoft; those names are used only to
describe compatibility.
