"""Word fayldan test savollarini o'qib, reyting tizimiga yuboruvchi vosita.

Ish tartibi: fayl(lar) yuklanadi -> /api/parse savollarni qaytaradi -> foydalanuvchi
ko'rib chiqadi va tuzatadi -> /api/send ularni reyting tizimi API'siga bo'laklab yuboradi.
"""

from __future__ import annotations

import json
import os
import re
import secrets
from typing import List, Optional

import httpx
from dotenv import load_dotenv
from fastapi import Depends, FastAPI, File, HTTPException, Query, Request, UploadFile, status
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, Field

from docx_parser import SLOTS, UNKNOWN_FORMULA, DocxError, parse_docx

load_dotenv()

# Reyting tizimi API'si: https://<domen>/api/public
API_BASE_URL = os.getenv("API_BASE_URL", "").rstrip("/")
API_KEY = os.getenv("API_KEY", "")
VERIFY_SSL = os.getenv("VERIFY_SSL", "true").lower() not in ("0", "false", "no")
# Har so'rovga qo'shiladigan cookie (masalan "usenew=1" - yangi tizim cookie orqali ochiladigan bosqichda).
API_COOKIE = os.getenv("API_COOKIE", "")
# Sahifani parol bilan yopish (bo'sh bo'lsa ochiq).
APP_USER = os.getenv("APP_USER", "admin")
APP_PASSWORD = os.getenv("APP_PASSWORD", "")

# Bitta so'rovda yuboriladigan savollar soni va hajmi (rasmlar base64 holida katta bo'ladi).
CHUNK_QUESTIONS = int(os.getenv("CHUNK_QUESTIONS", "40"))
CHUNK_BYTES = int(os.getenv("CHUNK_BYTES", str(4 * 1024 * 1024)))
MAX_FILE_BYTES = 30 * 1024 * 1024

app = FastAPI(title="Word -> savollar")
templates = Jinja2Templates(directory=os.path.join(os.path.dirname(__file__), "templates"))
basic = HTTPBasic(auto_error=False)


def guard(credentials: Optional[HTTPBasicCredentials] = Depends(basic)) -> None:
    """APP_PASSWORD berilgan bo'lsa, HTTP Basic orqali tekshiradi."""
    if not APP_PASSWORD:
        return
    valid = credentials is not None and secrets.compare_digest(
        credentials.username.encode(), APP_USER.encode()
    ) and secrets.compare_digest(credentials.password.encode(), APP_PASSWORD.encode())
    if not valid:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Login yoki parol noto'g'ri",
            headers={"WWW-Authenticate": "Basic"},
        )


def _headers() -> dict:
    headers = {"Accept": "application/json"}
    if API_KEY:
        headers["X-API-KEY"] = API_KEY
    if API_COOKIE:
        headers["Cookie"] = API_COOKIE
    return headers


def _client(timeout: float) -> httpx.AsyncClient:
    return httpx.AsyncClient(timeout=timeout, verify=VERIFY_SSL, headers=_headers())


def _api_error(response: httpx.Response) -> str:
    """Reyting API xato javobidan tushunarli matn chiqaradi."""
    if response.status_code == 401:
        return "API kaliti noto'g'ri yoki reyting tizimida sozlanmagan (API_KEY)."
    try:
        body = response.json()
    except ValueError:
        return f"Reyting tizimi xatosi ({response.status_code})."
    if isinstance(body, dict):
        if body.get("error"):
            return str(body["error"])
        if isinstance(body.get("errors"), dict):
            first = next(iter(body["errors"].values()), None)
            if first:
                return str(first[0] if isinstance(first, list) else first)
        if body.get("message"):
            return str(body["message"])
    return f"Reyting tizimi xatosi ({response.status_code})."


async def _proxy_get(path: str, params: dict | None = None) -> JSONResponse:
    if not API_BASE_URL:
        return JSONResponse({"success": False, "error": "API_BASE_URL sozlanmagan (.env)."}, status_code=500)
    try:
        async with _client(15.0) as client:
            response = await client.get(f"{API_BASE_URL}/{path}", params=params)
    except httpx.HTTPError as error:
        return JSONResponse(
            {"success": False, "error": f"Reyting tizimiga ulanib bo'lmadi: {error}"}, status_code=502
        )
    if response.status_code >= 400:
        return JSONResponse({"success": False, "error": _api_error(response)}, status_code=response.status_code)
    try:
        return JSONResponse(response.json())
    except ValueError:
        return JSONResponse({"success": False, "error": "Reyting tizimi JSON qaytarmadi."}, status_code=502)


@app.get("/", response_class=HTMLResponse, dependencies=[Depends(guard)])
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html", {"unknown_formula": UNKNOWN_FORMULA})


@app.get("/api/languages", dependencies=[Depends(guard)])
async def languages():
    return await _proxy_get("languages")


@app.get("/api/tests", dependencies=[Depends(guard)])
async def tests():
    return await _proxy_get("tests")


@app.get("/api/subjects", dependencies=[Depends(guard)])
async def subjects(
    language: str = "uz",
    test_id: Optional[int] = None,
    class_ids: Optional[List[int]] = Query(None),
):
    """Fanlar. `class_ids` berilsa, faqat o'sha sinflarda (CRM dars jadvali bo'yicha) o'tiladiganlari."""
    params: dict = {"language": language}
    if test_id:
        params["test_id"] = test_id
    if class_ids:
        params["class_ids[]"] = class_ids
    return await _proxy_get("subjects", params)


@app.get("/api/classes", dependencies=[Depends(guard)])
async def classes(level: int):
    """Sinf darajasidagi faol sinflar (fanlarni saralash uchun)."""
    return await _proxy_get("classes", {"level": level})


@app.get("/api/grades", dependencies=[Depends(guard)])
async def grades(test_id: Optional[int] = None):
    return await _proxy_get("grades", {"test_id": test_id} if test_id else None)


@app.post("/api/parse", dependencies=[Depends(guard)])
async def parse(files: List[UploadFile] = File(...)):
    """Word fayllarni o'qib, savollarni ko'rib chiqish uchun qaytaradi (hech narsa yubormaydi)."""
    result = []
    for upload in files:
        entry = {"name": upload.filename or "fayl.docx", "author": "", "questions": [], "warnings": [], "error": None}
        content = await upload.read()
        if len(content) > MAX_FILE_BYTES:
            entry["error"] = "Fayl juda katta (30 MB dan oshmasin)."
        else:
            try:
                entry.update(parse_docx(content))
            except DocxError as error:
                entry["error"] = str(error)
            except Exception as error:  # noqa: BLE001 - bitta fayl qolganlarini to'xtatmasin
                entry["error"] = f"Faylni o'qishda kutilmagan xatolik: {error}"
        result.append(entry)
    return {"success": True, "files": result}


class Part(BaseModel):
    text: str = ""
    image: Optional[str] = None


class QuestionIn(BaseModel):
    row: int = 0
    question: Part
    correct: Part
    wrong1: Part = Part()
    wrong2: Part = Part()
    wrong3: Part = Part()


class SendRequest(BaseModel):
    test_id: int
    language: str
    subject_id: int
    grade_id: int
    author: Optional[str] = None
    # Bo'sh bo'lsa savollar butun sinf darajasiga, aks holda faqat shu sinflarga tushadi.
    class_ids: List[int] = Field(default_factory=list)
    questions: List[QuestionIn] = Field(min_length=1)


def _chunks(items: list[dict]) -> list[list[dict]]:
    """Savollarni son va hajm chegarasi bo'yicha bo'laklaydi."""
    chunks: list[list[dict]] = []
    current: list[dict] = []
    size = 0
    for item in items:
        weight = len(json.dumps(item, ensure_ascii=False))
        if current and (len(current) >= CHUNK_QUESTIONS or size + weight > CHUNK_BYTES):
            chunks.append(current)
            current, size = [], 0
        current.append(item)
        size += weight
    if current:
        chunks.append(current)
    return chunks


@app.post("/api/send", dependencies=[Depends(guard)])
async def send(payload: SendRequest):
    """Ko'rib chiqilgan savollarni reyting tizimiga bo'laklab yuboradi."""
    if not API_BASE_URL:
        return JSONResponse({"success": False, "error": "API_BASE_URL sozlanmagan (.env)."}, status_code=500)

    items = []
    for question in payload.questions:
        data = question.model_dump()
        row = data.pop("row")
        for slot in SLOTS:
            data[slot]["text"] = data[slot]["text"].strip()
        items.append((row, data))

    totals = {"saved": 0, "skipped": 0, "linked": 0, "failed": 0, "errors": []}
    fatal = None

    async with _client(180.0) as client:
        for chunk in _chunks([data for _, data in items]):
            rows = [row for row, _ in items[: len(chunk)]]
            body = {
                "test_id": payload.test_id,
                "language": payload.language,
                "subject_id": payload.subject_id,
                "grade_id": payload.grade_id,
                "author": (payload.author or "").strip() or None,
                "class_ids": sorted(set(payload.class_ids)),
                "questions": chunk,
            }
            try:
                response = await client.post(f"{API_BASE_URL}/questions", json=body)
            except httpx.HTTPError as error:
                fatal = f"Reyting tizimiga ulanib bo'lmadi: {error}"
                break
            if response.status_code >= 400:
                fatal = _api_error(response)
                break
            try:
                answer = response.json()
            except ValueError:
                fatal = "Reyting tizimi JSON qaytarmadi."
                break

            # Eski reyting tizimi class_ids ni jimgina e'tiborsiz qoldiradi va savollar butun
            # darajaga tushib ketadi. Yangisi linked_count qaytaradi — bo'lmasa to'xtaymiz.
            if payload.class_ids and "linked_count" not in answer:
                fatal = ("Reyting tizimi sinf bo'yicha yuklashni qo'llamaydi (eski versiya): "
                         "savollar butun sinf darajasiga tushgan bo'lishi mumkin. Yuborish to'xtatildi.")
                items = items[len(chunk):]
                totals["saved"] += int(answer.get("saved_count", 0))
                break

            items = items[len(chunk):]
            totals["saved"] += int(answer.get("saved_count", 0))
            totals["skipped"] += int(answer.get("skipped_count", 0))
            totals["linked"] += int(answer.get("linked_count", 0))
            totals["failed"] += int(answer.get("failed_count", 0))
            for message in answer.get("errors") or []:
                totals["errors"].append(_with_row(str(message), rows))

    return {
        "success": fatal is None and totals["failed"] == 0,
        "saved": totals["saved"],
        "skipped": totals["skipped"],
        "linked": totals["linked"],
        "class_ids": sorted(set(payload.class_ids)),
        "failed": totals["failed"],
        "unsent": len(items),
        "errors": totals["errors"],
        "error": fatal,
    }


def _with_row(message: str, rows: list[int]) -> str:
    """API xabaridagi bo'lak ichidagi tartib raqamini Word jadvalidagi qator raqamiga almashtiradi."""
    match = re.match(r"^(\d+)-savol: (.*)$", message, re.S)
    if not match:
        return message
    index = int(match.group(1)) - 1
    if 0 <= index < len(rows) and rows[index]:
        return f"{rows[index]}-qator: {match.group(2)}"
    return message
