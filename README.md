# VBA-PdfParser

A unified VBA PDF text extraction wrapper that chains together four specialised helpers: [PdfTXT](https://github.com/rafael-yml/VBA-PdfTXT), [PdfWRT](https://github.com/rafael-yml/VBA-PdfWRT), [WinOCR](https://github.com/rafael-yml/VBA-WinOCR), and [WdCOM](https://github.com/rafael-yml/VBA-WdCOM).

Automatically falls back through extraction tiers until text is recovered.

---

## Why this exists

Each helper handles a different PDF scenario:

- **PdfTXT** – pure VBA text extraction from PDF content streams (no OCR, no external tools)
- **PdfWRT** – renders PDF pages to PNG using Windows.Data.Pdf (WinRT)
- **WinOCR** – extracts text from images via Windows.Media.Ocr
- **WdCOM** – Word COM automation as a last-resort fallback

PdfParser orchestrates them in a tiered pipeline so you get text with minimal boilerplate.

---

## Installation

Import these 5 files into your VBA project:

1. `PdfParser.cls` – main wrapper
2. `helpers/PdfTXT/PdfTXT.cls`
3. `helpers/PdfWRT/PdfWRT.cls`
4. `helpers/WinOCR/WinOCR.cls`
5. `helpers/WdCOM/WdCOM.cls`

No references to set, no external dependencies. Requires Windows 10+ for WinRT PDF rendering.

---

## Usage

### Simple extraction

```vb
Dim reader As New PdfParser
Dim sText As String

sText = reader.ExtractText("C:\docs\invoice.pdf")

Debug.Print reader.LastTier    ' "PdfTXT", "OCR", or "WdCOM"
Debug.Print reader.LastStatus  ' 0 = success, see PdfTXT constants
```

### Force Word fallback

```vb
sText = reader.ExtractText("C:\docs\complex.pdf", True)
```

The optional second parameter enables the WdCOM tier as a final fallback.

### Per-page extraction

```vb
Dim pages As Collection
Set pages = reader.ExtractPages("C:\docs\invoice.pdf")

Dim page As Variant
For Each page In pages
    Debug.Print page
Next
```

### Tuning individual helpers

```vb
reader.TXT.LineTolerance = 12      ' adjust text positioning
reader.WRT.DefaultWidth = 3508     ' render at A4 300dpi
reader.OCR.Language = "en-US"      ' set OCR language
```

---

## Tier flow

```
ExtractText(path)
    │
    ├─► PdfTXT.ExtractText
    │       │
    │       └─ success → return text, LastTier = "PdfTXT"
    │
    ├─► PdfWRT.RenderPDFToBytes → WinOCR.BytesToText (per page)
    │       │
    │       └─ success → return combined text, LastTier = "OCR"
    │
    └─► (if useWord = True) → WdCOM.ExtractText
            │
            └─ success → return text, LastTier = "WdCOM"
```

If all tiers fail, an empty string is returned.

---

## Dependencies

| Helper | What it does |
|---|---|
| [VBA-PdfTXT](https://github.com/rafael-yml/VBA-PdfTXT) | Pure VBA PDF text extraction |
| [VBA-PdfWRT](https://github.com/rafael-yml/VBA-PdfWRT) | PDF → PNG via WinRT |
| [VBA-WinOCR](https://github.com/rafael-yml/VBA-WinOCR) | Image → text via Windows OCR |
| [VBA-WdCOM](https://github.com/rafael-yml/VBA-WdCOM) | Word COM automation fallback |

All helpers are included as git submodules in the `helpers/` directory.

---

## Properties

#### `LastTier` → `String`

Which extraction tier succeeded. Values: `"PdfTXT"`, `"OCR"`, `"WdCOM"`, or `""`.

#### `LastStatus` → `Long`

Status code from PdfTXT. See PdfTXT constants: `PDFTXT_OK` (0), `PDFTXT_NO_TEXT` (1), `PDFTXT_NO_CMAP` (2), `PDFTXT_GARBLED` (3), `PDFTXT_FAIL` (4).

#### `TXT` → `PdfTXT`

Access to the underlying PdfTXT instance for tuning.

#### `WRT` → `PdfWRT`

Access to the underlying PdfWRT instance for tuning.

#### `OCR` → `WinOCR`

Access to the underlying WinOCR instance for tuning.

#### `WRD` → `WdCOM`

Access to the underlying WdCOM instance for tuning.

---

## Error handling

```vb
Dim reader As New PdfParser
Dim text As String

text = reader.ExtractText("C:\path\to\file.pdf")

If reader.LastTier = "" Then
    Debug.Print "All tiers failed - file may be encrypted or corrupt"
Else
    Debug.Print "Extracted via: " & reader.LastTier
End If
```

---

## License

MIT License. See [LICENSE](LICENSE) for details.

---

## Credits

- [VBA-PdfTXT](https://github.com/rafael-yml/VBA-PdfTXT)
- [VBA-PdfWRT](https://github.com/rafael-yml/VBA-PdfWRT)
- [VBA-WinOCR](https://github.com/rafael-yml/VBA-WinOCR)
- [VBA-WdCOM](https://github.com/rafael-yml/VBA-WdCOM)

Copyright © 2026, [rafael-yml](https://rafael-yml.lovable.app/)