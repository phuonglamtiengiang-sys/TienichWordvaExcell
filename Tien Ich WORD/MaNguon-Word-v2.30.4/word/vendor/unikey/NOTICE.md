# UniKey / UVConverter attribution and provenance

Upstream: UniKey UVConverter 1.1.3b, VnConv library.
Author: Pham Kim Long. Copyright (C) 1998–2002 Pham Kim Long.
Official source index: https://www.unikey.org/source.html
Original distribution URL: https://prdownloads.sourceforge.net/unikey/uvconv-src-1.1.3b.zip

The SourceForge URL returned an HTML page during this build. The source ZIP was obtained from this public archive mirror:
https://raw.githubusercontent.com/AutoLISP-Just-Simple/AJS_VnConvert/master/uvconv-src-1.1.3b.zip
This mirror is not represented as an official UniKey repository. The ZIP's upstream source headers identify VnConv, its author and GPL-2.0-or-later license. No claim of verification against an official archive checksum is made. Exact input hashes are recorded in SHA256SUMS.txt.

Unchanged upstream files are preserved under `upstream/uvconv/` and in the source ZIP. They include the upstream GPL text, AUTHORS, CREDITS and original build files.

Adaptation for Tiện ích Word (2026):

- `generate_unikey_tables.py` extracts all 213 entries in the TCVN3, VNI-WIN and Unicode tables from `vnconv/data.cpp`, without table changes. `UniKeyTables.generated.cs` is generated output.
- `UniKeyDecoder.cs` ports the needed SingleByteCharset and DoubleByteCharset constructor/nextInput logic and Unicode output mapping from `vnconv/charset.cpp` to C#.
- The Word adapter maps displayed Windows-1252 characters back to legacy bytes and returns source positions for edits that preserve surrounding formatting.
- Characters outside the legacy byte encoding are preserved. Orphan VNI tone bytes produce an error before editing instead of being silently discarded as upstream's INVALID_STD_CHAR handling would do.
- The existing explicit TCVN3 uppercase-font option remains a Word UI option after conversion.

This is a C# adaptation, not an official UniKey binary, and does not require UniKey to be installed or running. It implements TCVN3/VNI-WIN input to Unicode only, not all UVConverter encodings.

License: GNU GPL version 2 or (at your option) any later version. The Word add-in source/build files are supplied with the binary release under GPL-2.0-or-later. No warranty. See LICENSE-GPL-2.0.txt and the original license headers. The author's supplied donation image is a separate resource and is not third-party UniKey code.
