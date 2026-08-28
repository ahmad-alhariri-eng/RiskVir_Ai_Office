import json
import logging
from fastapi import APIRouter
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel
from typing import List, Dict, Any, Optional

from core.prompt_builder import get_system_prompt

logger = logging.getLogger("officeai")

router = APIRouter()


class ChatRequest(BaseModel):
    message: str
    context: Dict[str, Any]
    history: List[Dict[str, str]] = []


# ─── Context Summarizer ───────────────────────────────────────────

def _build_context_str(ctx: Dict[str, Any]) -> str:
    """
    Converts the rich frontend context object into a structured,
    LLM-readable string that gives the model precise spreadsheet/doc
    awareness so it can target correct cells, ranges, and structures.
    """
    app = ctx.get("app", "unknown")
    sel = ctx.get("selection")

    if sel is None:
        return f"App: {app}\nNo document context available."

    lines = [f"App: {app}"]

    # ── Excel context ────────────────────────────────────────────
    if isinstance(sel, dict) and "selectedRange" in sel:
        sr    = sel.get("selectedRange", {})
        used  = sel.get("usedRange")
        sheets = sel.get("worksheets", [])
        tables = sel.get("tables", [])
        named  = sel.get("namedRanges", [])
        active = sel.get("activeSheet", "")

        lines.append(f"Active Sheet: {active}")

        if sheets:
            sheet_names = [s["name"] for s in sheets]
            lines.append(f"All Sheets: {', '.join(sheet_names)}")

        # Selected range
        addr = sr.get("address", "")
        values = sr.get("values", [])
        formulas = sr.get("formulas", [])
        rows = sr.get("rowCount", 0)
        cols = sr.get("columnCount", 0)

        if addr:
            lines.append(f"\nSelected Range: {addr} ({rows}×{cols})")

        if values:
            # Show up to first 5 rows to keep context tight
            preview = values[:5]
            lines.append("Selected Values (first 5 rows):")
            for row in preview:
                lines.append("  " + " | ".join(str(v) for v in row))

        if formulas:
            # Only show cells that actually have formulas
            formula_cells = [
                f"  {sr['address']} row {ri+1}: {row}"
                for ri, row in enumerate(formulas[:5])
                if any(str(v).startswith("=") for v in row)
            ]
            if formula_cells:
                lines.append("Formulas in selection:")
                lines.extend(formula_cells[:5])

        # Used range overview
        if used:
            u_addr   = used.get("address", "")
            u_rows   = used.get("rowCount", 0)
            u_cols   = used.get("columnCount", 0)
            headers  = used.get("firstRow", [[]])[0] if used.get("firstRow") else []
            last_row = used.get("lastRow", [[]])[0] if used.get("lastRow") else []

            lines.append(f"\nUsed Range: {u_addr} ({u_rows} rows × {u_cols} cols)")
            if headers:
                lines.append(f"Column Headers: {' | '.join(str(h) for h in headers)}")
            if last_row:
                lines.append(f"Last Data Row:  {' | '.join(str(v) for v in last_row)}")

        # Tables
        if tables:
            lines.append("\nTables in workbook:")
            for t in tables:
                headers_str = ", ".join(t.get("headers", []))
                lines.append(f"  [{t['name']}] range={t['range']} headers=({headers_str})")

        # Named ranges
        if named:
            lines.append("Named Ranges:")
            for n in named:
                lines.append(f"  {n['name']} → {n['address']}")

    # ── Word context ─────────────────────────────────────────────
    elif isinstance(sel, str):
        if sel.strip():
            # Truncate very long selections to 1500 chars
            preview = sel[:1500] + ("…[truncated]" if len(sel) > 1500 else "")
            lines.append(f"\nSelected Text:\n{preview}")
        else:
            lines.append("No text selected.")

    # ── PowerPoint context ───────────────────────────────────────
    elif isinstance(sel, list):
        # Raw list of text parts from slide shapes
        text = " ".join(str(p) for p in sel if p)
        if text.strip():
            lines.append(f"\nCurrent Slide Text:\n{text[:1000]}")
        else:
            lines.append("Current slide has no text.")

    else:
        lines.append(f"Context data: {json.dumps(sel, ensure_ascii=False)[:500]}")

    return "\n".join(lines)


# ─── Endpoints ────────────────────────────────────────────────────

@router.get("/health")
def health_check():
    from main import engine
    return {
        "status": "ok",
        "model_loaded": engine is not None,
        "provider": "deepseek-api",
    }


@router.post("/chat")
async def chat_endpoint(req: ChatRequest):
    from main import engine
    if not engine:
        return JSONResponse(
            status_code=503,
            content={"error": "Model is not loaded. Please download the model and restart."}
        )

    app_name = req.context.get("app", "unknown")
    system_prompt = get_system_prompt(app_name)
    context_str   = _build_context_str(req.context)

    # Build message list: system → history → user
    messages: List[Dict[str, str]] = [{"role": "system", "content": system_prompt}]
    messages.extend(req.history)
    messages.append({
        "role": "user",
        "content": f"[DOCUMENT CONTEXT]\n{context_str}\n\n[USER REQUEST]\n{req.message}"
    })

    logger.info(
        f"Chat request | app={app_name} | "
        f"history={len(req.history)} turns | "
        f"context_len={len(context_str)} chars"
    )

    def event_generator():
        try:
            for chunk in engine.stream_chat(messages):
                yield f"data: {json.dumps({'token': chunk}, ensure_ascii=False)}\n\n"
            yield "data: [DONE]\n\n"
        except Exception as e:
            logger.error(f"Streaming error: {e}")
            yield f"data: {json.dumps({'error': str(e)})}\n\n"
            yield "data: [DONE]\n\n"

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        }
    )
