#!/usr/bin/env python3
"""
pipeline_atualizar.py — Pipeline completo do Indicador Ansell:

  1. Extrai do portal Brudam os relatorios de Ansell e Hercules (Playwright)
  2. Consolida os dois em uma planilha unica (salva tambem na pasta
     "Analise Ansell" do OneDrive, como no processo manual)
  3. Atualiza index.html com os dados novos
  4. Faz commit + push (so se algo realmente mudou)

Credenciais do portal vem de PORTAL_USER / PORTAL_PASS (variaveis de
ambiente) — nunca ficam no codigo. O envio de e-mail de notificacao usa
EMAIL_USER / EMAIL_PASS (conta Gmail + senha de app — o Office365 da
PortoEx bloqueia autenticacao SMTP por padrao), tambem via variavel de
ambiente. O e-mail sempre vai para EMAIL_DESTINO (PortoEx), so o
remetente e o Gmail.

Uso:
  set PORTAL_USER=seu.usuario
  set PORTAL_PASS=sua.senha
  set EMAIL_USER=seu.email@gmail.com
  set EMAIL_PASS=sua.senha.de.app.do.gmail
  python pipeline_atualizar.py
"""

import os
import smtplib
import subprocess
import sys
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

import atualizar_dashboard as ad
import extrair_portal as ep
from config import CONFIG

REPO_DIR = Path(__file__).parent
ONEDRIVE_ANALISE_ANSELL = Path(CONFIG["onedrive_consolidado"])
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_DESTINO = CONFIG["email_destino"]


LOG_FILE = REPO_DIR / "pipeline.log"
LOCK_FILE = REPO_DIR / "pipeline.lock"
LOCK_MAX_IDADE_MIN = 30  # acima disso, considera trava travada de uma execucao anterior que morreu


def log(msg):
    print(msg, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} — {msg}\n")
    except Exception:
        pass  # logging nunca deve derrubar o pipeline


def extrair_arquivos(usuario, senha, data_ini, data_fim, pasta_tmp: Path) -> dict:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.set_default_timeout(60000)

        try:
            ep.login(page, usuario, senha)

            arquivos = {}
            for cliente in ep.CLIENTES:
                arquivos[cliente] = ep.extrair_cliente(page, cliente, data_ini, data_fim, pasta_tmp)
        except Exception:
            diag_dir = pasta_tmp / "diagnostico"
            diag_dir.mkdir(parents=True, exist_ok=True)
            try:
                page.screenshot(path=str(diag_dir / "falha.png"), full_page=True)
                (diag_dir / "falha.html").write_text(page.content(), encoding="utf-8")
                log(f"Diagnostico salvo em: {diag_dir}")
            except Exception as diag_err:
                log(f"Nao foi possivel salvar diagnostico: {diag_err}")
            raise
        finally:
            browser.close()
    return arquivos


def consolidar(arquivos: dict, destino: Path) -> Path:
    log("\nConsolidando planilhas...")
    dfs = [pd.read_excel(caminho, sheet_name="Brudam") for caminho in arquivos.values()]
    consolidado = pd.concat(dfs, ignore_index=True)
    destino.parent.mkdir(parents=True, exist_ok=True)
    consolidado.to_excel(destino, sheet_name="Brudam", index=False)
    log(f"Consolidado ({len(consolidado)} linhas) salvo em: {destino}")
    return destino


def _sem_timestamp(raw):
    meta = dict(raw.get("meta", {}))
    meta.pop("gerado_em", None)
    return {"meta": meta, "rows": raw.get("rows")}


def atualizar_html(xlsx_consolidado: Path) -> bool:
    """Atualiza o index.html. O timestamp 'gerado_em' e sempre renovado
    (para refletir a ultima vez que a rotina rodou), mas o retorno indica
    se os dados de frete em si (fora do timestamp) realmente mudaram."""
    log("\nAtualizando index.html...")
    df = pd.read_excel(xlsx_consolidado, sheet_name="Brudam")
    rows = ad.build_rows(df)
    raw = ad.build_raw(rows)
    log(f"Periodo: {raw['meta']['meses'][0]} a {raw['meta']['meses'][-1]} | {len(rows)} linhas")

    import json

    index_path = REPO_DIR / "index.html"
    index_content = index_path.read_text(encoding="utf-8")

    raw_antigo = ad.read_json_blob(index_content, "const RAW = ")
    dados_mudaram = not (raw_antigo and _sem_timestamp(raw_antigo) == _sem_timestamp(raw))
    log("Dados de frete mudaram." if dados_mudaram else "Dados de frete iguais aos ja publicados.")

    raw_json = json.dumps(raw, ensure_ascii=False)
    index_content = ad.replace_json_blob(index_content, "const RAW = ", raw_json)
    index_path.write_text(index_content, encoding="utf-8")

    log("index.html atualizado.")
    return dados_mudaram


def limpar_desktop_ini_do_git():
    """O Google Drive (esta pasta e sincronizada) recria arquivos
    desktop.ini dentro de .git/refs/, corrompendo as referencias do
    Git e travando push/fetch silenciosamente. Remove antes de mexer
    no git."""
    git_dir = REPO_DIR / ".git"
    removidos = list(git_dir.rglob("desktop.ini"))
    for p in removidos:
        p.unlink(missing_ok=True)
    if removidos:
        log(f"Removidos {len(removidos)} desktop.ini de dentro do .git (Google Drive).")


def commit_e_push(dados_mudaram: bool):
    limpar_desktop_ini_do_git()
    log("\nVerificando alteracoes no git...")
    status = subprocess.run(
        ["git", "status", "--porcelain", "index.html"],
        cwd=REPO_DIR, capture_output=True, text=True,
    )
    log(f"git status --porcelain index.html -> rc={status.returncode} stdout={status.stdout!r} stderr={status.stderr!r}")
    if status.returncode != 0:
        log("git status falhou -- abortando commit desta rodada.")
        return
    if not status.stdout.strip():
        log("Nada para commitar (index.html identico ao publicado).")
        return

    subprocess.run(["git", "add", "index.html"], cwd=REPO_DIR, check=True)
    sufixo = "com dados novos" if dados_mudaram else "sem dados novos, so verificacao"
    mensagem = f"Atualizacao automatica ({sufixo}) — {date.today().isoformat()}"
    subprocess.run(["git", "commit", "-m", mensagem], cwd=REPO_DIR, check=True)
    subprocess.run(["git", "push"], cwd=REPO_DIR, check=True)
    log("Commit e push feitos com sucesso.")


def enviar_email(assunto, corpo_html):
    remetente = os.environ.get("EMAIL_USER")
    senha = os.environ.get("EMAIL_PASS")
    if not remetente or not senha:
        log("EMAIL_USER/EMAIL_PASS nao definidos — notificacao por e-mail pulada.")
        return

    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = formataddr(("Portoex x Ansell", remetente))
    msg["To"] = EMAIL_DESTINO
    msg.set_content(corpo_html, subtype="html")

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(remetente, senha)
            smtp.send_message(msg)
        log(f"E-mail de notificacao enviado para {EMAIL_DESTINO}.")
    except Exception as e:
        log(f"Falha ao enviar e-mail de notificacao: {e}")


def main():
    usuario = os.environ.get("PORTAL_USER")
    senha = os.environ.get("PORTAL_PASS")
    if not usuario or not senha:
        log("Defina as variaveis de ambiente PORTAL_USER e PORTAL_PASS antes de rodar.")
        sys.exit(1)

    hoje = date.today()
    data_ini = date(hoje.year, 1, 1).strftime("%d/%m/%Y")
    data_fim = (hoje - timedelta(days=1)).strftime("%d/%m/%Y")
    log(f"Periodo: {data_ini} ate {data_fim}")

    pasta_tmp = REPO_DIR / "downloads_tmp"
    arquivos = extrair_arquivos(usuario, senha, data_ini, data_fim, pasta_tmp)

    mes_abrev = ad.MES_ABREV[hoje.month]
    prefixo = CONFIG.get("prefixo_consolidado", "Base")
    nome_consolidado = f"{prefixo}_Jan_{mes_abrev}.xlsx"
    # No seu PC, guarda o consolidado no OneDrive (como no processo manual).
    # Na nuvem (GitHub Actions) essa pasta nao existe — usa uma pasta local.
    if ONEDRIVE_ANALISE_ANSELL.parent.exists():
        destino_consolidado = ONEDRIVE_ANALISE_ANSELL / nome_consolidado
    else:
        destino_consolidado = REPO_DIR / "downloads_tmp" / nome_consolidado
    consolidado_path = consolidar(arquivos, destino_consolidado)

    dados_mudaram = atualizar_html(consolidado_path)
    commit_e_push(dados_mudaram)

    link = CONFIG["link_dashboard"]
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    situacao = "<b><u>COM ALTERAÇÃO DE DADOS</u></b>" if dados_mudaram else "<b><u>SEM ALTERAÇÃO DE DADOS</u></b>"
    corpo_html = (
        f"<p>A rotina rodou normalmente em {agora}, {situacao}.</p>"
        f"<p>Acesse: <a href='{link}'>{link}</a></p>"
    )
    enviar_email("Indicador Ansell - Atualizado com Sucesso", corpo_html)

    log("\nPipeline concluido.")


def adquirir_lock():
    """Evita duas execucoes reais simultaneas (ex.: Task Scheduler disparando
    de novo com uma anterior ainda rodando) fazendo login duplicado no portal.
    Trava obsoleta (processo anterior que travou/morreu) e destravada sozinha
    apos LOCK_MAX_IDADE_MIN."""
    if LOCK_FILE.exists():
        idade_min = (datetime.now().timestamp() - LOCK_FILE.stat().st_mtime) / 60
        if idade_min < LOCK_MAX_IDADE_MIN:
            log(f"Ja existe uma execucao em andamento (trava com {idade_min:.1f} min) — abortando esta chamada.")
            return False
        log(f"Trava antiga encontrada ({idade_min:.1f} min) — considerando travada, removendo e seguindo.")
    LOCK_FILE.write_text(datetime.now().isoformat(), encoding="utf-8")
    return True


def liberar_lock():
    LOCK_FILE.unlink(missing_ok=True)


if __name__ == "__main__":
    log(f"\n{'='*60}\nInicio do pipeline: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    if not adquirir_lock():
        sys.exit(0)
    try:
        main()
    except Exception:
        import traceback
        log("ERRO NAO TRATADO:\n" + traceback.format_exc())
        raise
    finally:
        liberar_lock()
