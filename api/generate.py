import os
import sys
import json
import io
import time
from http.server import BaseHTTPRequestHandler
from pptx import Presentation

from google import genai
from google.genai import types

# Vercel loads this module as "api.generate", so the sibling "src" package is
# not on sys.path by default. Add this file's directory (the api/ folder) so the
# absolute "src.*" imports below resolve at runtime.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.theme import get_theme, list_themes
from src.slide_builder import SLIDE_BUILDERS

PUBLIC_DIR = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "public")
)

_CONTENT_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".svg": "image/svg+xml",
    ".ico": "image/x-icon",
}


def _is_retryable_model_error(exc):
    message = str(exc).upper()
    return any(code in message for code in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED"))


def _generate_content_with_retry(client, *, retries=2, **kwargs):
    last_error = None
    for attempt in range(retries + 1):
        try:
            return client.models.generate_content(**kwargs)
        except Exception as exc:
            last_error = exc
            if attempt >= retries or not _is_retryable_model_error(exc):
                raise
            time.sleep(1.5 * (attempt + 1))
    raise last_error


class handler(BaseHTTPRequestHandler):

    def do_GET(self):
        # On Vercel this function is the single catch-all entrypoint, so it must
        # also serve the static frontend bundled from the public/ directory.
        path = self.path.split("?", 1)[0]
        if path in ("", "/", "/api/generate"):
            path = "/index.html"

        file_path = os.path.normpath(os.path.join(PUBLIC_DIR, path.lstrip("/")))
        # Guard against path traversal outside public/.
        if not file_path.startswith(PUBLIC_DIR):
            self.send_error_response(404, "Not Found")
            return
        if not os.path.isfile(file_path) and not path.startswith("/api/"):
            fallback_path = os.path.join(PUBLIC_DIR, "index.html")
            if os.path.isfile(fallback_path):
                file_path = fallback_path
        if not os.path.isfile(file_path):
            self.send_error_response(404, "Not Found")
            return

        _, ext = os.path.splitext(file_path)
        content_type = _CONTENT_TYPES.get(ext.lower(), "application/octet-stream")
        with open(file_path, "rb") as f:
            body = f.read()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        try:
            # 1. リクエストの読み取り
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            req_body = json.loads(post_data.decode('utf-8'))
            
            user_prompt = req_body.get('prompt', '')
            requested_theme = req_body.get('theme', 'auto')
            slide_count = req_body.get('slide_count', 8)
            
            if not user_prompt:
                self.send_error_response(400, "Prompt is required")
                return

            # API Key チェック
            api_key = os.environ.get("GEMINI_API_KEY")
            if not api_key or api_key == "your_api_key_here":
                self.send_error_response(500, "GEMINI_API_KEY is not configured on the server.")
                return
                
            client = genai.Client(api_key=api_key)
            
            # 2. Gemini 3.5 Flash でJSON構成案を作成
            # Vercelのタイムアウト制限(60s)を考慮し、画像生成は主要な3枚までに制限する指示を出す。
            system_instruction = f"""
            あなたは世界屈指のパワーポイントデザイナーです。ユーザーの指示に基づいて、プレゼンの構成をJSONで出力してください。
            使用できるスライドタイプ: title, content, image_full, two_column, key_message, ending
            利用可能なテーマ: {', '.join(list_themes())}
            
            【重要ルール - Vercelタイムアウト対策】
            画像生成の時間は限られているため、画像(imageフィールド)を設定するのは、「title」スライドと、「image_full」または特に重要な「content」スライドの最大3枚までにしてください。
            それ以外のスライドには image フィールドを含めないでください（美しいグラデーション背景が自動適用されます）。
            """
            
            # 期待するJSONスキーマの定義 (簡略化)
            schema = {
                "type": "object",
                "properties": {
                    "theme": {"type": "string", "description": "テーマ名 (midnight, elegance, corporate, nature, creative, warm, minimal)"},
                    "slides": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "type": {"type": "string"},
                                "title": {"type": "string"},
                                "subtitle": {"type": "string"},
                                "caption": {"type": "string"},
                                "message": {"type": "string"},
                                "sub_message": {"type": "string"},
                                "bullets": {
                                    "type": "array",
                                    "items": {
                                        "type": "object",
                                        "properties": {
                                            "title": {"type": "string"},
                                            "desc": {"type": "string"}
                                        }
                                    }
                                },
                                "left": {"type": "object", "properties": {"text": {"type": "array", "items": {"type": "string"}}}},
                                "right": {"type": "object", "properties": {"text": {"type": "array", "items": {"type": "string"}}}},
                                "image_prompt": {"type": "string", "description": "画像が必要な場合のみ、英語で詳細な画像生成プロンプトを記述。最大3スライドまで。"}
                            },
                            "required": ["type"]
                        }
                    }
                },
                "required": ["theme", "slides"]
            }

            response_json = _generate_content_with_retry(
                client,
                model='gemini-3.5-flash',
                contents=f"目的・内容:\n{user_prompt}\n\n希望スライド数: 約{slide_count}枚\n指定テーマ: {requested_theme}",
                config=types.GenerateContentConfig(
                    system_instruction=system_instruction,
                    response_mime_type="application/json",
                    response_schema=schema,
                    temperature=0.7,
                ),
            )
            
            presentation_config = json.loads(response_json.text)
            
            # テーマの上書き処理
            if requested_theme != 'auto' and requested_theme in list_themes():
                presentation_config['theme'] = requested_theme
                
            theme_obj = get_theme(presentation_config.get("theme", "minimal"))
            
            # 3. Gemini 3.1 Flash Image (Nano Banana 2) で画像生成
            # /tmp ディレクトリに保存 (Vercel Serverless Functionで唯一書き込み可能な場所)
            tmp_dir = "/tmp/images"
            os.makedirs(tmp_dir, exist_ok=True)
            
            for i, slide_data in enumerate(presentation_config.get("slides", [])):
                img_prompt = slide_data.get("image_prompt")
                if img_prompt:
                    # 画像生成
                    try:
                        img_response = _generate_content_with_retry(
                            client,
                            model='gemini-3.1-flash-image',
                            contents=img_prompt,
                            config=types.GenerateContentConfig(
                                temperature=0.7,
                            )
                        )
                        # google-genai の画像出力は response.candidates[0].content.parts[0].inline_data.data (bytes) で取得
                        # もしくは generated_image 等の専用プロパティがある。
                        # ドキュメントによれば、通常は bytes データを取得して保存する。
                        # もしエラーが出た場合はログを出して無視し、グラデーション背景にフォールバックさせる。
                        if img_response.candidates and img_response.candidates[0].content.parts:
                            part = img_response.candidates[0].content.parts[0]
                            if part.inline_data:
                                img_bytes = part.inline_data.data
                                file_path = os.path.join(tmp_dir, f"slide_{i}.png")
                                with open(file_path, "wb") as f:
                                    f.write(img_bytes)
                                slide_data["image"] = file_path
                    except Exception as e:
                        print(f"Image generation failed for slide {i}: {str(e)}")
                        # 失敗した場合はimageフィールドをセットしない（グラデーションフォールバック）
            
            # 4. PowerPoint ビルド
            prs = Presentation()
            # スライドサイズを16:9に設定
            prs.slide_width = 12192000  # 13.33 inches
            prs.slide_height = 6858000  # 7.5 inches
            
            for i, slide_data in enumerate(presentation_config.get("slides", [])):
                slide_type = slide_data.get("type")
                if slide_type in SLIDE_BUILDERS:
                    builder_func = SLIDE_BUILDERS[slide_type]
                    # title と ending 以外はページ番号を付与
                    page_num = 0 if slide_type in ["title", "ending"] else i + 1
                    if slide_type in ["title", "ending"]:
                        builder_func(prs, slide_data, theme_obj)
                    else:
                        builder_func(prs, slide_data, theme_obj, page_num)
                        
            # メモリバッファに保存
            pptx_stream = io.BytesIO()
            prs.save(pptx_stream)
            pptx_bytes = pptx_stream.getvalue()
            
            # 5. クライアントに返却
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
            self.send_header('Content-Disposition', 'attachment; filename="AI_Presentation.pptx"')
            self.send_header('Content-Length', str(len(pptx_bytes)))
            self.end_headers()
            
            self.wfile.write(pptx_bytes)

        except Exception as e:
            import traceback
            error_msg = f"Internal Server Error: {str(e)}\n{traceback.format_exc()}"
            print(error_msg)
            if _is_retryable_model_error(e):
                self.send_error_response(
                    503,
                    "The AI model is temporarily busy. Please try again in a moment.",
                )
            else:
                self.send_error_response(500, str(e))
            
    def send_error_response(self, code, message):
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps({"error": message}).encode('utf-8'))
