"""Diagnostico de acesso ao portal Brudam a partir da nuvem (sem credenciais):
mostra codigo HTTP, titulo e um trecho do texto da pagina, pra saber se o
formulario de login carrega ou se o acesso e barrado."""
from playwright.sync_api import sync_playwright

URL = "https://azportoex.brudam.com.br/"

with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    pg = b.new_context().new_page()
    resp = pg.goto(URL, wait_until="load", timeout=60000)
    pg.wait_for_timeout(5000)
    texto = " ".join(pg.inner_text("body").split())[:300] if pg.query_selector("body") else ""
    n_user = pg.get_by_placeholder("Usuário").count()
    n_inputs = pg.locator("input").count()
    print(f"status={resp.status if resp else None} url={pg.url} titulo={pg.title()!r} "
          f"campo_usuario={n_user} inputs={n_inputs} texto={texto!r}")
    b.close()
