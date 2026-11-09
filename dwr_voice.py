from pathlib import Path
from datetime import datetime, timedelta, timezone
import json
import os
import time
import logging
import zipfile

import gerar_dwr

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("America/Sao_Paulo")
except Exception:
    # Windows sem banco de fusos (tzdata): o Brasil não tem horário de verão desde 2019.
    TZ = timezone(timedelta(hours=-3))
BASE = Path(__file__).resolve().parent
CONFIG_FILE = BASE / "config.json"
STATE_FILE = BASE / "estado.json"


class AguardandoBase(Exception):
    """Falta o documento base (ou o DWR está aberto): o ditado fica na fila."""


def read_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception as exc:
            raise RuntimeError(f"Não foi possível ler config.json: {exc}")
    return {}


def _home_dirs():
    """Subpastas da pasta do usuário (ex.: C:\\Users\\nome), sem diferenciar maiúsculas."""
    try:
        return [p for p in Path.home().iterdir() if p.is_dir()]
    except OSError:
        return []


def find_icloud_drive(config):
    configured = (config.get("icloud_drive") or "").strip()
    if configured:
        path = Path(os.path.expandvars(configured)).expanduser()
        if path.exists():
            return path
        raise RuntimeError(f"A pasta iCloud configurada não existe: {path}")

    home = Path.home()
    candidates = [
        os.environ.get("iCloudDrive"),
        home / "iCloudDrive",
        home / "iCloud Drive",
        home / "iCloudDrive (Archive)",
    ]
    # Qualquer pasta cujo nome comece com "icloud" (varia conforme a versão do iCloud).
    candidates += [p for p in _home_dirs() if p.name.lower().startswith("icloud")]
    # Pasta que já contém DWR_Entrada (inclusive dentro do OneDrive, se o iCloud foi movido).
    candidates += [p for p in _home_dirs() if (p / "DWR_Entrada").exists()]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return Path(candidate)
    achadas = ", ".join(p.name for p in _home_dirs()) or "(nenhuma)"
    raise RuntimeError(
        "Não encontrei a pasta do iCloud Drive.\n"
        f"Pasta do usuário: {home}\nSubpastas encontradas: {achadas}\n"
        'Abra o config.json e preencha "icloud_drive" com o caminho completo da pasta '
        'iCloud Drive, usando barras duplas, ex.: "C:\\\\Users\\\\golle\\\\iCloudDrive".'
    )


def find_onedrive_dwr(config):
    configured = (config.get("onedrive_dwr") or "").strip()
    if configured:
        path = Path(os.path.expandvars(configured)).expanduser()
        if path.exists():
            return path
        raise RuntimeError(f"A pasta OneDrive/DWR configurada não existe: {path}")

    candidates = [
        os.environ.get("OneDrive"),
        os.environ.get("OneDriveCommercial"),
        os.environ.get("OneDriveConsumer"),
    ]
    candidates += [p for p in _home_dirs() if p.name.lower().startswith("onedrive")]
    for candidate in candidates:
        if candidate and (Path(candidate) / "DWR").exists():
            return Path(candidate) / "DWR"
    raise RuntimeError(
        "Não encontrei OneDrive/DWR. Confirme que o OneDrive está conectado "
        "e que a pasta DWR existe. Se necessário, preencha \"onedrive_dwr\" no config.json "
        "com o caminho completo da pasta DWR."
    )


config = read_config()
ICLOUD_ROOT = find_icloud_drive(config)
ONEDRIVE_DWR = find_onedrive_dwr(config)
INBOX_ROOT = ICLOUD_ROOT / "DWR_Entrada"
INBOX = INBOX_ROOT / "fila"
DONE = INBOX_ROOT / "feito"
ERRORS = INBOX_ROOT / "erros"
EXAMPLE = INBOX_ROOT / "exemplo"
REPORTS = ONEDRIVE_DWR / "relatorios"
LOGS = BASE / "dwr_voice.log"
# Formato da data escrita DENTRO do DWR (o nome do arquivo usa sempre dd-mm-aaaa).
DATE_FORMAT = config.get("formato_data") or "%m/%d/%Y"

for folder in (INBOX, DONE, ERRORS, EXAMPLE, REPORTS):
    folder.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    filename=LOGS,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8",
)


def report_path(day):
    # Windows não permite barras em nomes de arquivos.
    return REPORTS / f"DWR_{day.strftime('%d-%m-%Y')}.docx"


# ---------------------------------------------------------------- documento base

def find_base():
    """Documento base do serviço: exemplo/base.docx (ou o único .docx da pasta)."""
    preferred = EXAMPLE / "base.docx"
    if preferred.exists():
        return preferred
    docs = sorted(p for p in EXAMPLE.glob("*.docx") if not p.name.startswith("~$"))
    if len(docs) == 1:
        return docs[0]
    if not docs:
        raise AguardandoBase(
            f"Coloque o DWR base do serviço (cabeçalho + tabelas + assinaturas) em: {EXAMPLE}"
        )
    raise AguardandoBase(
        f"Há mais de um .docx em {EXAMPLE}. Deixe só um, ou chame-o de base.docx."
    )


def base_number(base):
    with zipfile.ZipFile(base) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    cells = gerar_dwr.CELL.findall(gerar_dwr.TBL.findall(xml)[0])
    return int(gerar_dwr.texto(cells[1]).strip())


# ---------------------------------------------------------------- estado do dia

def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    base = find_base()
    return {"numero": base_number(base) - 1, "data": None, "atividades": [], "acoes": []}


def save_state(state):
    tmp = STATE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, STATE_FILE)


def roll_day(state, day):
    """Devolve o estado de `day`: em um novo dia o número sobe 1 e os registros zeram."""
    state = json.loads(json.dumps(state))
    key = day.isoformat()
    if state.get("data") != key:
        state.update(numero=state["numero"] + 1, data=key, atividades=[], acoes=[])
    return state


def write_report(day, state):
    """Gera o DWR completo (base + data/número + tabelas) e substitui o arquivo do dia."""
    base = find_base()
    path = report_path(day)
    tmp = path.with_name(path.name + ".tmp")
    try:
        gerar_dwr.gerar(
            base, tmp, state["numero"], day.strftime(DATE_FORMAT),
            state["atividades"], state["acoes"],
        )
        os.replace(tmp, path)
    except PermissionError:
        raise AguardandoBase("O DWR do dia está aberto no Word; feche-o para continuar.")
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
    return path


def ensure_report(day):
    """Se o DWR do dia não existe, cria um novo a partir do base (tabelas vazias)."""
    old = load_state()
    state = roll_day(old, day)
    path = report_path(day)
    if not path.exists():
        write_report(day, state)
        logging.info("Relatório do dia criado: %s (DWR nº %02d)", path, state["numero"])
        print(f"Relatório criado: {path.name}")
    if state != old:
        save_state(state)
    return path


def create_or_update_report(day, category, description):
    state = roll_day(load_state(), day)
    state["atividades" if category == "ACT" else "acoes"].append(description)
    path = write_report(day, state)
    save_state(state)          # só depois de o Word ter sido gravado com sucesso
    return path


# ---------------------------------------------------------------- fila

def wait_for_complete_file(file_path):
    previous_size = -1
    stable_count = 0
    for _ in range(30):
        try:
            size = file_path.stat().st_size
            if size > 0 and size == previous_size:
                stable_count += 1
                if stable_count >= 2:
                    return
            else:
                stable_count = 0
            previous_size = size
        except FileNotFoundError:
            pass
        time.sleep(1)
    raise RuntimeError("O arquivo não terminou de sincronizar ou está indisponível.")


def process_file(file_path):
    name = file_path.name.upper()
    if name.startswith("ACT_"):
        category = "ACT"
    elif name.startswith("REQ_"):
        category = "REQ"
    else:
        raise ValueError("Use nomes iniciados por ACT_ ou REQ_.")

    wait_for_complete_file(file_path)
    description = file_path.read_text(encoding="utf-8-sig").strip()
    if not description:
        raise ValueError("O arquivo está vazio.")

    day = datetime.now(TZ).date()
    result = create_or_update_report(day, category, description)
    destination = DONE / file_path.name
    counter = 1
    while destination.exists():
        destination = DONE / f"{file_path.stem}_{counter}{file_path.suffix}"
        counter += 1
    file_path.replace(destination)
    logging.info("Registro %s incluído em %s", file_path.name, result)
    return result


def main():
    print("DWR Voice ativo.")
    print(f"Entrada iCloud: {INBOX}")
    print(f"Documento base: {EXAMPLE}")
    print(f"Relatórios OneDrive: {REPORTS}")
    print("ACT_*.txt = atividades | REQ_*.txt = ações/solicitações")
    logging.info("DWR Voice iniciado. Entrada=%s | Saída=%s", INBOX, REPORTS)

    last_warning = None

    def warn(message):
        nonlocal last_warning
        if message != last_warning:      # evita repetir o aviso a cada 2 segundos
            print(f"AVISO: {message}")
            last_warning = message

    while True:
        # Também verifica a cada ciclo (cobre virada do dia e relatório apagado).
        try:
            ensure_report(datetime.now(TZ).date())
            last_warning = None
        except AguardandoBase as exc:
            warn(str(exc))
        except Exception as exc:
            logging.exception("Falha ao verificar/criar o relatório do dia")
            print(f"ERRO ao criar o relatório do dia: {exc}")

        for file_path in sorted(INBOX.glob("*.txt")):
            try:
                result = process_file(file_path)
                print(f"Processado: {file_path.name} -> {result.name}")
            except AguardandoBase as exc:
                warn(str(exc))           # o ditado continua na fila e entra depois
                break
            except Exception as exc:
                logging.exception("Falha ao processar %s", file_path)
                try:
                    target = ERRORS / file_path.name
                    counter = 1
                    while target.exists():
                        target = ERRORS / f"{file_path.stem}_{counter}{file_path.suffix}"
                        counter += 1
                    file_path.replace(target)
                except Exception:
                    pass
                print(f"ERRO em {file_path.name}: {exc}")
        time.sleep(2)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Não foi possível iniciar o DWR Voice: {exc}")
        print("Confira config.json e os caminhos do iCloud Drive e OneDrive.")
        input("Pressione Enter para fechar...")
