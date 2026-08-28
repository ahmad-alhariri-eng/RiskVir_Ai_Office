# OfficeAI Pro

A local AI assistant that runs directly inside **Microsoft Word, Excel, and PowerPoint** as a Task Pane Add-in — powered by a local LLM (no cloud, no API keys required).

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│  Microsoft Office (Word / Excel / PowerPoint)               │
│  ┌───────────────────────────────────────────────────────┐  │
│  │  Task Pane Add-in  (React + TypeScript, Vite)         │  │
│  │  office-addin/src/App.tsx                              │  │
│  │  office-addin/src/utils/  (readers + actions)         │  │
│  └──────────────────────┬────────────────────────────────┘  │
└─────────────────────────│───────────────────────────────────┘
                          │ HTTPS SSE  /chat
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  Backend  (FastAPI + Python)                                │
│  backend/main.py         — app config, CORS, routing        │
│  backend/api/routes/     — /health, /chat endpoints         │
│  backend/core/           — LLMEngine, PromptBuilder         │
│  backend/services/       — extensible service layer         │
└──────────────────────────┬──────────────────────────────────┘
                           │
                           ▼
                 models/*.gguf  (local LLM via llama-cpp-python)

System Tray:  launcher/launcher.py  (pystray — Start/Stop backend)
```

---

## Quick Start

### Prerequisites
- Python 3.10+
- Node.js 18+
- ~3 GB disk space for the AI model

### 1. Install everything
```bat
scripts\setup.bat
```
This installs Python deps, downloads the Phi-4-mini model (~2.5 GB), and runs `npm install`.

### 2. Start services
```bat
scripts\start.bat
```
Opens two windows — backend on `http(s)://127.0.0.1:8000` and Vite dev server on `http://localhost:5173`.

### 3. Load the Add-in in Office
1. Open Word / Excel / PowerPoint
2. Go to **Insert → Add-ins → Upload My Add-in**
3. Select `office-addin/manifest.xml`
4. The OfficeAI Pro pane will appear on the right

---

## Configuration

Copy `.env.example` to `.env` and edit as needed:

```ini
# Use a custom model (optional — auto-detected from models/ if not set)
OFFICEAI_MODEL_PATH=C:\path\to\your\model.gguf

# Server
OFFICEAI_HOST=127.0.0.1
OFFICEAI_PORT=8000

# LLM tuning
OFFICEAI_N_THREADS=8
OFFICEAI_N_CTX=4096
OFFICEAI_MAX_TOKENS=16384
```

---

## Project Structure

```
microsoft/
├── .env.example             ← copy to .env
├── .gitignore
├── OfficeAI_Backend.spec    ← PyInstaller spec
├── OfficeAI_Launcher.spec   ← PyInstaller spec
│
├── backend/
│   ├── main.py              ← FastAPI app entry point
│   ├── requirements.txt
│   ├── core/
│   │   ├── llm_engine.py    ← llama-cpp-python wrapper
│   │   └── prompt_builder.py← System prompts (Word/Excel/PPT)
│   ├── api/
│   │   └── routes/
│   │       └── chat.py      ← /health + /chat endpoints
│   └── services/            ← extensible service layer
│
├── office-addin/
│   ├── manifest.xml         ← Office Add-in manifest
│   ├── src/
│   │   ├── App.tsx          ← Main UI component
│   │   ├── types/actions.ts ← TypeScript action types
│   │   └── utils/
│   │       ├── excelActions.ts / excelReader.ts
│   │       ├── wordActions.ts  / wordReader.ts
│   │       └── pptActions.ts   / pptReader.ts
│   └── package.json
│
├── launcher/
│   └── launcher.py          ← System tray controller
│
├── models/
│   └── *.gguf               ← Local LLM model (not in git)
│
└── scripts/
    ├── setup.bat            ← First-time setup
    ├── start.bat            ← Start both services
    └── build_exe.bat        ← Package as standalone EXE
```

---

## Supported Office Commands

### Word (`WordMind`)
| Command | Description |
|---------|-------------|
| `insertText` | Insert text at cursor |
| `insertHtml` | Insert HTML (tables, formatted text) |
| `replaceSelection` | Replace selected text |
| `applyStyle` | Apply Word style (Heading 1, Body Text…) |
| `insertPageBreak` | Insert page break |
| `findAndReplace` | Find & replace across document |
| `setDocumentProperty` | Set title, author, subject, custom props |

### Excel (`DataForge`)
| Command | Description |
|---------|-------------|
| `insertData` | Insert 2D data array |
| `addChart` | Create chart (bar/column/line/pie/scatter…) |
| `insertFormula` | Insert formula into specific cell |
| `formatRange` | Format cells (bold, color, font, alignment…) |
| `addConditionalFormat` | Color scale / data bar / icon set |
| `addSheet` | Add new worksheet |
| `setColumnWidth` | Set column widths |

### PowerPoint (`DeckMind`)
| Command | Description |
|---------|-------------|
| `addSlide` | Add slide with title & bullet points |
| `addSlideWithTable` | Add slide with styled data table |
| `addSlideWithChart` | Add slide with chart description |
| `addSpeakerNotes` | Add speaker notes to a slide |
| `formatSlide` | Format slide fonts & background |

---

## Build Standalone EXEs

```bat
scripts\build_exe.bat
```

Produces:
- `dist\OfficeAI_Backend.exe` — self-contained FastAPI server
- `dist\OfficeAI_Launcher.exe` — system tray controller

> **Note:** The model file (`models/*.gguf`) is NOT bundled. Copy it alongside the EXEs on the target machine.

---

## Using a Different Model

The backend auto-detects any `.gguf` file in the `models/` directory. To use a specific model:

```ini
# .env
OFFICEAI_MODEL_PATH=C:\Users\YourName\.lmstudio\models\your-model.gguf
```

Compatible with any GGUF model supported by `llama-cpp-python` (Llama, Mistral, Phi, Qwen, Gemma…).

---

## Development

```bat
# Backend only
cd backend
python main.py

# Frontend only
cd office-addin
npm run dev

# Type-check frontend
cd office-addin
npx tsc --noEmit
```
