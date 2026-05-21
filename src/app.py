import webview
from pathlib import Path

from backend.webview_api import WebViewApi


if __name__ == "__main__":
    base_dir = Path(__file__).resolve().parent

    login_file = base_dir / "web" / "pages" / "login.html"

    api = WebViewApi(base_dir=base_dir)

    window = webview.create_window(
        "Caíque Castro - Central",
        login_file.as_uri(),
        fullscreen=False,
        js_api=api,  # expõe a API para o JS (window.pywebview.api)
        width=1200,
        height=750,
    )
    api.attach_window(window)

    webview.start(debug=False)
