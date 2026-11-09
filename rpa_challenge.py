import argparse
import csv
import logging
import sys
import time
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

URL = "https://rpachallenge.com/"
EXCEL_URL = "https://rpachallenge.com/assets/downloadFiles/challenge.xlsx"

BASE = Path(__file__).resolve().parent
INPUT = BASE / "input"
OUTPUT = BASE / "output"
LOGS = BASE / "logs"

RUN_ID = datetime.now().strftime("%Y%m%d_%H%M%S")

log = logging.getLogger("rpa")

JS_PREENCHER = """
(dados) => {
    const norm = s => (s || '').replace(/\\s+/g, ' ').trim().toLowerCase();
    const form = document.querySelector('form');
    if (!form) return { ok: false, erro: 'Formulario nao encontrado' };

    const campos = {};
    form.querySelectorAll('label').forEach(lb => {
        const inp = lb.parentElement && lb.parentElement.querySelector('input');
        if (inp) campos[norm(lb.textContent)] = inp;
    });

    const setValue = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    const faltando = [];
    const divergente = [];

    for (const [nome, valor] of Object.entries(dados)) {
        const inp = campos[norm(nome)];
        if (!inp) { faltando.push(nome); continue; }
        setValue.call(inp, valor);
        inp.dispatchEvent(new Event('input', { bubbles: true }));
        inp.dispatchEvent(new Event('change', { bubbles: true }));
        if (inp.value !== valor) divergente.push(nome);
    }

    if (faltando.length || divergente.length) {
        let msg = '';
        if (faltando.length) msg += 'Campo nao encontrado: ' + faltando.join(', ') + '. ';
        if (divergente.length) msg += 'Valor nao aplicado: ' + divergente.join(', ');
        return { ok: false, erro: msg.trim() };
    }

    const btn = form.querySelector('input[type=submit]');
    if (!btn) return { ok: false, erro: 'Botao Submit nao encontrado' };
    btn.click();
    return { ok: true, qtd: Object.keys(dados).length };
}
"""


def config_log():
    LOGS.mkdir(exist_ok=True)
    arq = LOGS / f"execucao_{RUN_ID}.log"
    fmt = logging.Formatter("%(asctime)s.%(msecs)03d | %(levelname)s | %(message)s", "%d/%m/%Y %H:%M:%S")
    log.setLevel(logging.INFO)
    for h in (logging.FileHandler(arq, encoding="utf-8"), logging.StreamHandler(sys.stdout)):
        h.setFormatter(fmt)
        log.addHandler(h)
    return arq


def ler_excel(caminho):
    wb = load_workbook(caminho, read_only=True, data_only=True)
    linhas = wb.active.iter_rows(values_only=True)
    cabecalho = [str(c).strip() for c in next(linhas) if c is not None]
    registros = []
    for linha in linhas:
        if not linha or all(v is None for v in linha):
            continue
        registros.append({c: "" if v is None else str(v).strip() for c, v in zip(cabecalho, linha)})
    wb.close()
    return registros


def salvar_resultados(resultados):
    OUTPUT.mkdir(exist_ok=True)
    colunas = list(resultados[0].keys())

    xlsx = OUTPUT / f"resultado_{RUN_ID}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"
    ws.append(colunas)
    for r in resultados:
        ws.append([r[c] for c in colunas])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = min(60, max(len(str(c.value or "")) for c in col) + 2)
    wb.save(xlsx)

    arq_csv = OUTPUT / f"resultado_{RUN_ID}.csv"
    with open(arq_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=colunas, delimiter=";")
        w.writeheader()
        w.writerows(resultados)

    return xlsx, arq_csv


def baixar_excel(page):
    INPUT.mkdir(exist_ok=True)
    destino = INPUT / "challenge.xlsx"
    resp = page.request.get(EXCEL_URL)
    if not resp.ok:
        raise Exception(f"Erro ao baixar planilha (HTTP {resp.status})")
    destino.write_bytes(resp.body())
    return destino


def bloquear_recursos(route):
    if route.request.resource_type in ("image", "media"):
        route.abort()
    else:
        route.continue_()


def main(excel=None, headed=False):
    arq_log = config_log()
    inicio = time.perf_counter()
    resultados = []
    retorno = 0

    log.info("Iniciando execucao %s", RUN_ID)

    with sync_playwright() as p:
        log.info("Abrindo navegador")
        browser = p.chromium.launch(headless=not headed)
        context = browser.new_context(viewport={"width": 1366, "height": 768})
        context.route("**/*", bloquear_recursos)
        page = context.new_page()
        page.set_default_timeout(15000)

        try:
            if excel:
                log.info("Usando planilha local: %s", excel)
            else:
                log.info("Baixando planilha do site")
                excel = baixar_excel(page)

            registros = ler_excel(excel)
            if not registros:
                raise Exception("Planilha sem registros")
            log.info("%d registros lidos da planilha", len(registros))

            log.info("Acessando %s", URL)
            page.goto(URL, wait_until="domcontentloaded")
            page.locator("form input[type=submit]").wait_for()
            page.get_by_role("button", name="Start").click()
            log.info("Desafio iniciado")

            t = time.perf_counter()
            for i, reg in enumerate(registros, 1):
                dt = datetime.now().strftime("%d/%m/%Y %H:%M:%S.%f")[:-3]
                try:
                    res = page.evaluate(JS_PREENCHER, reg)
                    if res["ok"]:
                        status, obs = "OK", f"Registro enviado com sucesso ({res['qtd']} campos)"
                        log.info("Linha %d: OK", i)
                    else:
                        status, obs = "NOK", res["erro"]
                        log.warning("Linha %d: NOK - %s", i, obs)
                except PlaywrightError as e:
                    status, obs = "NOK", str(e).splitlines()[0]
                    log.error("Linha %d: NOK - %s", i, obs)

                resultados.append({"Linha": i, **reg, "Status": status, "Observacao": obs, "Data/Hora": dt})

            log.info("Formulario preenchido em %.0f ms", (time.perf_counter() - t) * 1000)

            msg = page.locator(".message2")
            msg.wait_for()
            log.info("Resultado: %s", " ".join(msg.inner_text().split()))

            OUTPUT.mkdir(exist_ok=True)
            print_final = OUTPUT / f"score_{RUN_ID}.png"
            page.screenshot(path=str(print_final), full_page=True)
            log.info("Print salvo em %s", print_final)

        except Exception as e:
            retorno = 1
            log.exception("Erro na execucao: %s", e)
            OUTPUT.mkdir(exist_ok=True)
            try:
                page.screenshot(path=str(OUTPUT / f"erro_{RUN_ID}.png"), full_page=True)
            except PlaywrightError:
                pass
        finally:
            context.close()
            browser.close()

    if resultados:
        xlsx, arq_csv = salvar_resultados(resultados)
        ok = sum(1 for r in resultados if r["Status"] == "OK")
        log.info("Total: %d OK / %d NOK", ok, len(resultados) - ok)
        log.info("Resultados salvos em %s e %s", xlsx, arq_csv)

    log.info("Fim da execucao (%.2f s) - log em %s", time.perf_counter() - inicio, arq_log)
    return retorno


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", type=Path)
    parser.add_argument("--headed", action="store_true")
    args = parser.parse_args()
    sys.exit(main(args.excel, args.headed))
