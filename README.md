# PO Translation System

An internal web application that translates Purchase Order (PO) Excel workbooks from English to **Tamil** and **Hindi**.

You upload a `.xlsx` file and get back one workbook that contains a Tamil copy and a Hindi copy of every sheet. Text is translated; PO numbers, material codes, part numbers, quantities, prices, dates and formulas stay exactly as they were, and the formatting of the original is kept.

- **Backend:** FastAPI (upload, translate, download endpoints)
- **Frontend:** Streamlit (one page: upload, translate, download)
- **Excel handling:** openpyxl
- **Translation:** Google Translate (free endpoint by default, or the paid Cloud Translation API)

## Quick start

Requires Python 3.10 or newer. Commands are for Windows PowerShell, run from the `po_translation_system` folder.

```powershell
# 1. Create a virtual environment and install dependencies
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt

# 2. Create your configuration (the defaults work without any changes)
Copy-Item .env.example .env

# 3. Start the API
python -m uvicorn app.main:app --port 8000
```

In a second terminal (activate the virtual environment again):

```powershell
python -m streamlit run frontend/streamlit_app.py   # Main command
```

Open [http://localhost:8501](http://localhost:8501), upload a PO workbook, click **Translate**, then download the result.

The API documentation (Swagger UI) is at [http://localhost:8000/docs](http://localhost:8000/docs).

## Configuration

Settings are read from environment variables or from a `.env` file in this folder. See [.env.example](.env.example).

| Variable                          | Default                                         | Purpose                                      |
| --------------------------------- | ----------------------------------------------- | -------------------------------------------- |
| `TRANSLATION_PROVIDER`          | `google_free`                                 | Which translator to use, see below           |
| `TRANSLATION_API_KEY`           | empty                                           | API key for the`google` provider           |
| `GOOGLE_TRANSLATE_API_ENDPOINT` | Cloud Translation v2 URL                        | Endpoint for the`google` provider          |
| `TRANSLATION_TIMEOUT`           | `30`                                          | Timeout in seconds for one Cloud API request |
| `MAX_FILE_SIZE`                 | `10485760`                                    | Largest accepted upload in bytes (10 MB)     |
| `ALLOWED_EXTENSIONS`            | `.xlsx`                                       | Accepted file extensions                     |
| `UPLOAD_DIR`                    | `uploads`                                     | Where uploaded files are stored              |
| `OUTPUT_DIR`                    | `outputs`                                     | Where translated files are stored            |
| `CORS_ORIGINS`                  | `http://localhost:8501,http://127.0.0.1:8501` | Origins allowed to call the API              |
| `API_BASE_URL`                  | `http://localhost:8000`                       | API address used by the Streamlit frontend   |

### Translation providers

| Provider        | API key    | Use it for                                                                                                           |
| --------------- | ---------- | -------------------------------------------------------------------------------------------------------------------- |
| `google_free` | Not needed | Getting started and light use. Uses the public Google Translate web endpoint through the`deep-translator` package. |
| `google`      | Required   | Regular or production use. Uses the Google Cloud Translation API v2.                                                 |
| `mock`        | Not needed | Tests and offline demos. Returns`[ta] text` / `[hi] text` instead of a real translation.                         |

Things to know before choosing:

- **`google_free` is rate limited.** Google allows about 5 requests per second and blocks an IP address for a while when that is exceeded. The application sends up to 100 cells per request and waits between requests to stay under the limit, but there is no service guarantee. If translation returns "temporarily unavailable" repeatedly, wait and retry, or switch to `google`.
- **Both Google providers send the text of your PO cells to Google.** Check that this is acceptable for your purchase order data. Codes, numbers and formulas that are protected (see below) are not sent.
- To use `google`: enable the Cloud Translation API in a Google Cloud project, create an API key, then set `TRANSLATION_PROVIDER=google` and `TRANSLATION_API_KEY=<your key>` in `.env`.

## API

| Method   | Path                         | Description                                                                    |
| -------- | ---------------------------- | ------------------------------------------------------------------------------ |
| `POST` | `/api/upload`              | Upload a`.xlsx` file (multipart form field `file`). Returns a `file_id`. |
| `POST` | `/api/translate/{file_id}` | Translate the uploaded file. Returns the output filename.                      |
| `GET`  | `/api/download/{file_id}`  | Download the translated workbook.                                              |
| `GET`  | `/health`                  | Health check, shows the active translation provider.                           |

Example with curl:

```bash
curl -F "file=@MyPO.xlsx" http://localhost:8000/api/upload
# {"file_id":"3f0c...","filename":"MyPO.xlsx","status":"uploaded"}

curl -X POST http://localhost:8000/api/translate/3f0c...
# {"status":"completed","languages":["Tamil","Hindi"],"output_filename":"MyPO_Translated.xlsx","translated_cells":42}

curl -OJ http://localhost:8000/api/download/3f0c...
```

Error responses:

| Status  | Meaning                                                                     |
| ------- | --------------------------------------------------------------------------- |
| `400` | Not an`.xlsx` file, unreadable or empty workbook, or nothing to translate |
| `404` | Unknown`file_id`, or download requested before translation                |
| `413` | File larger than`MAX_FILE_SIZE`                                           |
| `503` | The translation provider failed or is rate limiting                         |

## How translation works

1. The uploaded file is saved to `uploads/` as `<file_id>_<name>.xlsx` and is never modified. Work happens on a temporary copy.
2. Every text cell is examined. Protected values are skipped, the rest is collected as text to translate. Each distinct text is translated only once.
3. The output keeps every English sheet and adds one copy per language (`Tamil_<sheet>`, `Hindi_<sheet>`) with the translations written into the copies. Pictures, shapes, borders, merged cells and print setup are carried over unchanged. Sheets of the upload that are already in Tamil or Hindi are left out.
4. The result is saved to `outputs/` as `<file_id>_<name>_Translated.xlsx` and downloaded as `<name>_Translated.xlsx`.

### What is not translated

| Kind of value                                 | Examples                                                                         |
| --------------------------------------------- | -------------------------------------------------------------------------------- |
| Numbers, percentages, currency amounts        | `100`, `5,000.00`, `15%`, `$50.00`                                       |
| Dates                                         | `07/10/2026`, `2026-07-10`                                                   |
| Formulas                                      | `=B5*B6`                                                                       |
| PO, material, part, employee and vendor codes | `PO-10025`, `PO/2026/001`, `MAT-1001`, `JCB-12345`, `EMP001`, `V002` |
| Any single word that mixes letters and digits | `M8x40`, `X200`                                                              |
| URLs and email addresses                      | `https://example.com`, `purchase@example.com`                                |

A code inside a sentence is also preserved: `MAT-1001 Hydraulic Pump` becomes `MAT-1001 ஹைட்ராலிக் பம்ப்`.

As a safety check, if a translation comes back with different numbers than the English text (for example `30` turned into `3`), the English text is kept for that cell and a warning is logged.

The rules live in `ExcelService` in [app/services/excel_service.py](app/services/excel_service.py). To protect another code format, add a pattern there and list it in `_is_code`.

### What is kept from the original workbook

Cell styles (fonts, fills, borders, number formats, alignment), merged cells, column widths, row heights, frozen panes, filters, conditional formatting, data validation (dropdowns), comments, hyperlinks, page setup, headers and footers, and images such as a logo. Formulas that refer to another sheet are redirected to the copy of that sheet in the same language.

### Limitations

- Only `.xlsx` is supported (not `.xls`, `.xlsm` or `.csv`).
- Charts, text boxes, shapes and pivot tables are not carried over to the output.
- Text in headers, footers, comments and dropdown lists is copied but not translated.
- Text produced by a formula is not translated.
- Company names and addresses are sent to the translator and usually come back transliterated.
- Machine translation can be wrong. Have a Tamil or Hindi speaker review the output before it is sent to a supplier.
- Uploaded and translated files stay in `uploads/` and `outputs/` until someone deletes them. There is no automatic cleanup.
- The API has no login. Run it only on a trusted internal network.

## Project structure

```
po_translation_system/
├── app/
│   ├── main.py                     FastAPI application
│   ├── config.py                   Settings from environment / .env
│   ├── api/
│   │   ├── upload.py               POST /api/upload
│   │   ├── translate.py            POST /api/translate/{file_id}
│   │   └── download.py             GET  /api/download/{file_id}
│   ├── services/
│   │   ├── excel_service.py        Finds translatable cells, builds the output workbook
│   │   └── translation_service.py  Translation providers (google_free, google, mock)
│   ├── models/schemas.py           API response models
│   └── utils/                      Logging, validation, filename helpers
├── frontend/streamlit_app.py       Streamlit user interface
├── tests/                          pytest test suite
├── uploads/                        Uploaded files (created at runtime)
├── outputs/                        Translated files (created at runtime)
├── .env.example
└── requirements.txt
```

## Tests

```powershell
python -m pytest
```

The tests use the `mock` provider and temporary folders, so they need no network access and do not touch `uploads/` or `outputs/`.

To create a sample PO workbook to try the application with:

```powershell
python -c "from pathlib import Path; from app.services.excel_service import ExcelService; ExcelService().create_sample_po(Path('Sample_PO.xlsx'))"
```

## Troubleshooting

| Problem                                                              | Cause and fix                                                                                                                                                                                                                                        |
| -------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Frontend shows "Cannot reach the API"                                | The backend is not running, or`API_BASE_URL` points to the wrong address.                                                                                                                                                                          |
| "Translation service is temporarily unavailable"                     | With`google_free`: Google is rate limiting your IP address. Wait and retry, or switch to the `google` provider. With `google`: check `TRANSLATION_API_KEY` and that the Cloud Translation API is enabled. The API log shows the exact error. |
| "The uploaded workbook does not contain usable data for translation" | The workbook has no English text cells, only numbers, codes or formulas.                                                                                                                                                                             |
| The logo is missing in the output                                    | Pillow is not installed. Run`pip install -r requirements.txt`.                                                                                                                                                                                     |
| `ModuleNotFoundError: No module named 'app'`                       | Run the commands from the`po_translation_system` folder.                                                                                                                                                                                           |
