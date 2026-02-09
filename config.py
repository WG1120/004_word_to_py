"""프로젝트 설정 및 환경변수 로드."""

import os
from dotenv import load_dotenv

load_dotenv()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
DEFAULT_MODEL = "gpt-4o-mini"
TEMPERATURE = 0.2
MAX_RETRIES = 3
RETRY_BASE_DELAY = 5  # seconds
