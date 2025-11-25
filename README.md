# Hindi Dialogue Translator

A Python script to translate dialogue columns in DOCX, XLSX, or CSV files into Hinglish (Roman Hindi) using Gemini AI or offline Argos Translate.

## Features

- **Dual Engine Support**: Use Gemini AI (online) or Argos Translate (offline)
- **Multiple Formats**: Supports `.docx`, `.xlsx`, and `.csv` files
- **Hinglish Translation**: Kid-friendly, dub-friendly Hinglish translations with Maizen-style tone
- **Batch Processing**: Process multiple files at once
- **Detailed Logging**: Track progress with configurable log levels

## Installation

1. Clone this repository:
```bash
git clone <your-repo-url>
cd hindi-translator
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Single File Translation

**With Gemini (requires API key):**
```powershell
python translate_dialogues.py "input.docx" --engine gemini --api-key "YOUR_API_KEY" --model "gemini-2.5-flash" --delay-seconds 7 --output "output.hinglish.docx"
```

**With Argos (offline, no API key needed):**
```powershell
python translate_dialogues.py "input.docx" --engine argos --output "output.hinglish.docx"
```

### Batch Processing (Multiple Files)

```powershell
Get-ChildItem -File -Filter *.docx | ForEach-Object {
  python translate_dialogues.py "$($_.FullName)" --engine gemini --api-key "YOUR_API_KEY" --model "gemini-2.5-flash" --delay-seconds 7 --log-level INFO --output "$($_.DirectoryName)\$($_.BaseName).hinglish$($_.Extension)"
}
```

### Options

- `--engine {gemini,argos}`: Translation engine (default: `argos`)
- `--api-key`: Gemini API key (or set `GOOGLE_API_KEY` env var)
- `--model`: Gemini model name (default: `gemini-2.5-flash`)
- `--source-column`: Name of source column (default: `Dialogue`)
- `--target-column`: Name of target column (default: `Translation`)
- `--delay-seconds`: Delay between API requests (default: `7.0`)
- `--max-rows`: Limit translation to N rows (default: `0` = all)
- `--log-level`: Logging level (default: `INFO`)

### Getting a Gemini API Key

1. Visit [Google AI Studio](https://aistudio.google.com/app/apikey)
2. Sign in and create a new API key
3. Use it with `--api-key` or set `GOOGLE_API_KEY` environment variable

## File Structure

Your input files should have:
- A **Dialogue** column (or specify with `--source-column`)
- A **Translation** column (or specify with `--target-column`)

The script will fill empty cells in the Translation column.

## Output

The script creates a new file with `.hinglish` suffix (or as specified with `--output`), preserving the original file structure.

## All Rights Reserved by ANUSHREE

