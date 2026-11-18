# docx_api

Word (.docx) jadvalidan test savollarini o'qib, reyting tizimiga yuboruvchi vosita.

## Word fayl ko'rinishi

Savollar jadval ichida bo'ladi, har qator bitta savol:

| Savol | To'g'ri javob | Noto'g'ri javob | Noto'g'ri javob | Noto'g'ri javob | Muallif (ixtiyoriy) |
|---|---|---|---|---|---|

Sarlavha qatori bo'lsa o'zi tashlab ketiladi. Bo'sh qatorlar hisobga olinmaydi.

## Formulalar

Fayl qanday bo'lsa shundayligicha yuklanadi, formulalar LaTeX'ga o'girilib matn ichiga
`\( ... \)` ko'rinishida qo'yiladi:

- Word formula muharriri (Вставка → Уравнение);
- MathType va eski Equation 3.0 obyektlari;
- oddiy matn bilan yozilgan daraja va indekslar (x², H₂O).

O'girib bo'lmagan formula o'rniga `[FORMULA?]` qo'yiladi va savol qizil bilan belgilanadi —
uni ko'rib chiqish sahifasida qo'lda yozish kerak. Reyting tizimi formulalarni KaTeX bilan chizadi.

## O'rnatish va ishga tushirish

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp .env.example .env   # API_BASE_URL va API_KEY ni to'ldiring
.venv/bin/uvicorn main:app --port 8010
```

`API_KEY` reyting tizimining `.env` faylidagi `QUESTIONS_API_KEY` bilan bir xil bo'lishi kerak.

## Testlar

```bash
.venv/bin/pip install pytest
.venv/bin/pytest
```
