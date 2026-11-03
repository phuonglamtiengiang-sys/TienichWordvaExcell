# Building Tiện ích Word 2.30.4

License: the Word source code in this distribution is provided under GPL-2.0-or-later. See LICENSE-GPL-2.0.txt. UniKey's original notices are retained under vendor/unikey/. The donation PNG is an author-supplied resource, separate from UniKey code.

Extract the complete source archive into a folder. The top level must contain `word`, `tests` and `about`.

Requirements: Windows 64-bit, Python 3, Windows PowerShell, .NET Framework 4.x C# compiler at `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`. The add-in is built AnyCPU; only Word 64-bit has been tested.

Build from the source archive root:

```
python word/build_word.py
```

The build regenerates the C# tables from the preserved upstream UniKey `data.cpp`, validates Ribbon XML with the supplied OfficeDev schema, and compiles the add-in and installer into `output/Word-2.30.4`. It does not download or execute external source during the build.

Original icons are supplied as transparent PNG resources. To regenerate the artwork, install Pillow in your Python environment and run `python word/draw_ribbon_icons.py` on Windows. The script reads Segoe UI Semibold from Windows; no font files are distributed. All 24 icons have distinct original artwork and are covered by the Word source license. `tests/test_word_icons.ps1` verifies their Office picture wrappers and cache.

Verify every legacy table entry:

```
python tests/test_unikey_tables.py
```

Verify edits and Undo in your own test Word document (requires Word installed):

```
powershell -NoProfile -ExecutionPolicy Bypass -File tests/test_word_editing.ps1
```

Tests create and discard their own temporary document. They do not change Office security settings. Installation is a separate user action; close Word and use the generated installer. An existing installation is updated in the current user's LocalAppData folder.

The exact preserved UniKey archive/data/decoder hashes are in vendor/unikey/SHA256SUMS.txt. The C# decoder adaptation and its differences from upstream are documented in vendor/unikey/NOTICE.md.
