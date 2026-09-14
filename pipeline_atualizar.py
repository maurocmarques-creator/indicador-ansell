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

REPO_DIR = Path(__file__).parent
ONEDRIVE_ANALISE_ANSELL = Path(
    r"C:\Users\Mauro Cesar Marques\OneDrive - PORTOEXPRESS LOGISTICA LTDA\Analise Ansell"
)
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_DESTINO = "mauro.cesar@portoex.com.br"


def log(msg):
    print(msg, flush=True)


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


def atualizar_html(xlsx_consolidado: Path):
    log("\nAtualizando index.html...")
    df = pd.read_excel(xlsx_consolidado, sheet_name="Brudam")
    rows = ad.build_rows(df)
    raw = ad.build_raw(rows)
    log(f"Periodo: {raw['meta']['meses'][0]} a {raw['meta']['meses'][-1]} | {len(rows)} linhas")

    import json

    index_path = REPO_DIR / "index.html"
    index_content = index_path.read_text(encoding="utf-8")

    # Se os dados forem identicos aos ja publicados (fora do timestamp),
    # mantem o "gerado_em" antigo para nao gerar um commit so por causa
    # da hora em que o script rodou.
    raw_antigo = ad.read_json_blob(index_content, "const RAW = ")
    if raw_antigo and _sem_timestamp(raw_antigo) == _sem_timestamp(raw):
        raw["meta"]["gerado_em"] = raw_antigo["meta"].get("gerado_em")
        log("Dados iguais aos ja publicados — mantendo timestamp anterior.")

    raw_json = json.dumps(raw, ensure_ascii=False)
    index_content = ad.replace_json_blob(index_content, "const RAW = ", raw_json)
    index_path.write_text(index_content, encoding="utf-8")

    log("index.html atualizado.")


def commit_e_push():
    log("\nVerificando alteracoes no git...")
    status = subprocess.run(
        ["git", "status", "--porcelain", "index.html"],
        cwd=REPO_DIR, capture_output=True, text=True, check=True,
    )
    if not status.stdout.strip():
        log("Nenhuma alteracao nos dados — nada para commitar.")
        return False

    subprocess.run(["git", "add", "index.html"], cwd=REPO_DIR, check=True)
    mensagem = f"Atualizacao automatica dos dados ({date.today().isoformat()})"
    subprocess.run(["git", "commit", "-m", mensagem], cwd=REPO_DIR, check=True)
    subprocess.run(["git", "push"], cwd=REPO_DIR, check=True)
    log("Commit e push feitos com sucesso.")
    return True


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
    nome_consolidado = f"Base_Jan_{mes_abrev}.xlsx"
    # No seu PC, guarda o consolidado no OneDrive (como no processo manual).
    # Na nuvem (GitHub Actions) essa pasta nao existe — usa uma pasta local.
    if ONEDRIVE_ANALISE_ANSELL.parent.exists():
        destino_consolidado = ONEDRIVE_ANALISE_ANSELL / nome_consolidado
    else:
        destino_consolidado = REPO_DIR / "downloads_tmp" / nome_consolidado
    consolidado_path = consolidar(arquivos, destino_consolidado)

    atualizar_html(consolidado_path)
    houve_atualizacao = commit_e_push()

    link = "https://maurocmarques-creator.github.io/indicador-ansell/"
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    situacao = "<b><u>COM ALTERAÇÃO DE DADOS</u></b>" if houve_atualizacao else "<b><u>SEM ALTERAÇÃO DE DADOS</u></b>"
    corpo_html = (
        f"<p>A rotina rodou normalmente em {agora}, {situacao}.</p>"
        f"<p>Acesse: <a href='{link}'>{link}</a></p>"
    )
    enviar_email("Indicador Ansell - Atualizado com Sucesso", corpo_html)

    log("\nPipeline concluido.")


if __name__ == "__main__":
    main()
