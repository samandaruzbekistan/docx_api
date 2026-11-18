#!/bin/bash
# Word parser -> reyting tizimi (reyting.ideal-study.uz).
# API kaliti reyting tizimining deploy .env faylidan o'qiladi (ekranga chiqmaydi).
cd "$(dirname "$0")" || exit 1
ENV_FILE="${RATING_ENV_FILE:-/Users/mac/Projects/ideal/rating-deploy.env}"
export API_BASE_URL="http://reyting.ideal-study.uz/api/public"
export API_KEY="$(grep -E '^QUESTIONS_API_KEY=' "$ENV_FILE" | head -1 | cut -d= -f2- | tr -d '"\r')"
# Yangi tizim hali cookie orqali ochiladigan bosqich uchun; to'liq almashtirilgandan keyin ham zarar qilmaydi.
export API_COOKIE="usenew=1"
if [ -z "$API_KEY" ]; then echo "QUESTIONS_API_KEY topilmadi: $ENV_FILE"; exit 1; fi
echo "Parser: http://127.0.0.1:${PORT:-8765}  ->  $API_BASE_URL"
exec .venv/bin/uvicorn main:app --host 127.0.0.1 --port "${PORT:-8765}"
