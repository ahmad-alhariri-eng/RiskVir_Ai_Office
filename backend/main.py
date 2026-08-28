import os
import logging
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv

from core.llm_engine import LLMEngine
from api.routes.chat import router as chat_router

# Load .env from project root (parent of backend/)
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_ENV_PATH = os.path.join(_ROOT, '.env')
load_dotenv(_ENV_PATH, override=True)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("officeai")

engine = None



@asynccontextmanager
async def lifespan(app: FastAPI):
    global engine
    try:
        engine = LLMEngine()
        logger.info("DeepSeek engine initialized successfully.")
    except Exception as e:
        logger.error(f"Failed to initialize DeepSeek engine: {e}")
        engine = None
    yield


app = FastAPI(title="OfficeAI Pro Backend", lifespan=lifespan)

ALLOWED_ORIGINS = [
    "https://localhost:8000",
    "https://127.0.0.1:8000",
    "https://localhost:5173",   # Vite dev server
    "http://localhost:5173",    # Vite dev server (HTTP fallback)
    "http://localhost:8000",
    "http://127.0.0.1:8000",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Register API routes
app.include_router(chat_router)

# Serve static frontend build if available
STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "office-addin", "dist")
if os.path.exists(STATIC_DIR):
    app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    logger.info("Starting with HTTP (plain) — tunnel handles HTTPS")
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
