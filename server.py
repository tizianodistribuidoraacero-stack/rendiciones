from flask import Flask, request, send_file, send_from_directory, jsonify
from openpyxl import load_workbook
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import io
import os
import smtplib
import traceback

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)

TEMPLATE = os.path.join(BASE_DIR, "PLANILLA DE RENDICION v1.0.xlsx")
INDEX_HTML = os.path.join(BASE_DIR, "index.html")  # <- el HTML va acá, junto al server.py

RANGES = {
    "GIGANTE": (5, 7),
    "OBRAS": (9, 11),
    "LM": (13, 18),
}

# E=EFECTIVO, F=TRANSFERENCIA, H=CHEQUE, I=E-CHEQ, K=RETENCION, J=AJUSTE CENTAVOS
MEDIO_TO_COL = {
    "EFECTIVO": "E",
    "TRANSF.": "F",
    "CHEQUES": "H",
    "ECHEQ": "I",
    "RETENCIONES": "K",
    "AJUSTE": "J",
}

MAIL_CC = "nicolasd@distribuidoracero.com.ar"

def find_next_row(ws, start, end):
    for r in range(start, end + 1):
        if ws[f"D{r}"].value in (None, ""):
            return r
    return None

def add_number(ws, cell_addr, value):
    current = ws[cell_addr].value
    try:
        current_num = float(current) if current not in (None, "") else 0.0
    except Exception:
        current_num = 0.0
    ws[cell_addr].value = current_num + float(value)

def build_excel_from_clients(clients):
    if not os.path.exists(TEMPLATE):
        raise FileNotFoundError(f"No encuentro la plantilla: {TEMPLATE}")

    wb = load_workbook(TEMPLATE)
    ws = wb["Rendición"] if "Rendición" in wb.sheetnames else wb.active
    ws["C3"].value = datetime.now().strftime("%d/%m/%Y")

    for c in clients:
        modal = c.get("modal")
        cli = (c.get("cli") or "").strip().upper()
        items = c.get("items", [])

        if modal not in RANGES or not cli or not items:
            continue

        start, end = RANGES[modal]
        row = find_next_row(ws, start, end)
        if row is None:
            continue

        ws[f"D{row}"].value = cli

        for it in items:
            med = it.get("med")
            imp = it.get("imp")

            if med not in MEDIO_TO_COL:
                continue
            try:
                imp = float(imp)
            except Exception:
                continue

            col = MEDIO_TO_COL[med]
            add_number(ws, f"{col}{row}", imp)

    filename = f"rendicion-{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue(), filename

def clients_summary(clients):
    lines = []
    for c in clients:
        n = len(c.get("items", []))
        plural = "ítem" if n == 1 else "ítems"
        lines.append(f"- {c.get('modal')} · {c.get('cli')} ({n} {plural})")
    return "\n".join(lines)

def send_rendicion_email(to_addrs, cc_addrs, clients, excel_bytes, filename):
    smtp_host = os.environ.get("SMTP_HOST")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))
    smtp_user = os.environ.get("SMTP_USER")
    smtp_password = os.environ.get("SMTP_PASSWORD")
    mail_from = os.environ.get("MAIL_FROM", smtp_user)

    if not smtp_host or not smtp_user or not smtp_password:
        raise RuntimeError(
            "SMTP no configurado en el servidor. "
            "Configurá SMTP_HOST, SMTP_USER y SMTP_PASSWORD en Render."
        )

    fecha = datetime.now().strftime("%d/%m/%Y")
    subject = f"Rendición Distribuidora Acero - {fecha}"
    body = (
        "Adjunto planilla de rendición.\n\n"
        f"Clientes incluidos:\n{clients_summary(clients)}\n"
    )

    msg = MIMEMultipart()
    msg["From"] = mail_from
    msg["To"] = ", ".join(to_addrs)
    if cc_addrs:
        msg["Cc"] = ", ".join(cc_addrs)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    part = MIMEBase(
        "application",
        "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    part.set_payload(excel_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)

    recipients = list(dict.fromkeys([*to_addrs, *cc_addrs]))
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(mail_from, recipients, msg.as_string())

@app.get("/")
def home():
    if not os.path.exists(INDEX_HTML):
        # mensaje claro si el HTML no está donde debe
        return (
            "No encuentro index.html en la misma carpeta que server.py.\n"
            "Poné el archivo como: <carpeta>/index.html",
            404,
        )
    return send_from_directory(BASE_DIR, "index.html")

@app.get("/health")
def health():
    return "OK", 200

@app.post("/generar")
def generar():
    try:
        data = request.get_json(force=True, silent=True)
        if not data:
            return "JSON inválido", 400

        clients = data.get("clients", [])
        if not clients:
            return "Sin datos", 400

        excel_bytes, filename = build_excel_from_clients(clients)
        bio = io.BytesIO(excel_bytes)

        return send_file(
            bio,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception:
        # devolvemos el error real para que no quede "vacío"
        return "ERROR:\n" + traceback.format_exc(), 500

@app.post("/enviar-mail")
def enviar_mail():
    try:
        data = request.get_json(force=True, silent=True)
        if not data:
            return "JSON inválido", 400

        clients = data.get("clients", [])
        if not clients:
            return "Sin datos", 400

        cc_only = bool(data.get("cc_only", False))
        excel_bytes, filename = build_excel_from_clients(clients)

        if cc_only:
            to_addrs = [MAIL_CC]
            cc_addrs = []
        else:
            mail_to = (data.get("to") or os.environ.get("MAIL_TO") or MAIL_CC).strip()
            to_addrs = [mail_to]
            cc_addrs = [] if mail_to.lower() == MAIL_CC.lower() else [MAIL_CC]

        send_rendicion_email(to_addrs, cc_addrs, clients, excel_bytes, filename)
        return jsonify({"ok": True, "cc": MAIL_CC}), 200

    except Exception:
        return "ERROR:\n" + traceback.format_exc(), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run("0.0.0.0", port, debug=False)

