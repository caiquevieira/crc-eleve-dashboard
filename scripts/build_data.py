"""Gera site/data.json a partir da aba "Produtividade" da planilha CRC - ELEVE.

Somente campos operacionais são exportados. Nome, telefone, data de nascimento
e observações NUNCA saem deste script, porque o GitHub Pages é público.

Uso:
  python scripts/build_data.py --out site/data.json            # baixa do Drive (env SHEET_FILE_ID)
  python scripts/build_data.py --xlsx planilha.xlsx --out site/data.json
"""
from __future__ import annotations

import argparse
import datetime as dt
import io
import json
import os
import re
import sys
import unicodedata
import urllib.request

import openpyxl

SHEET = "Produtividade"

# Índices das colunas (A=0) na aba Produtividade
COL_DATA_AGEND = 0
COL_NOME = 1
COL_CAMPANHA = 3
COL_CRC = 4
COL_CLINICA = 6
COL_DATA_CONSULTA = 7
COL_HORA = 8
COL_CONFIRMACAO = 9
COL_MEIO = 10
COL_STATUS = 11
COL_FECHOU = 12
COL_ORCAMENTO = 13
COL_VALOR = 14

CLINICAS = {
    "PENH": "Penha", "PENHA": "Penha", "00PENHA": "Penha",
    "GRI": "Grimaldi", "GRIMALDI": "Grimaldi", "00GRIMALDI": "Grimaldi",
    "TEO": "Teotônio", "TEOTONIO": "Teotônio",
    "DIVA": "Vila Diva", "VILA DIVA": "Vila Diva",
    "PIRA": "Pira", "STAR": "Star", "ELEVE": "Eleve",
}

CAMPANHAS = {
    "ANIVERSARIO": "Aniversariantes", "ANIVERSARIANTES": "Aniversariantes",
    "INDICACAO": "Indicação", "LEAD NOVO": "Lead novo", "BASE FACE": "Base face",
    "FACEBOOK": "Facebook", "CAMPANHA": "Campanha", "IMPLANTE": "Implante",
    "LIMPEZA": "Limpeza", "LIMPEZA 250": "Limpeza 250", "PROTESE FLEX": "Prótese flex",
    "PROTOCOLO": "Protocolo", "VOUCHER 250": "Voucher 250", "VOUCHER 500": "Voucher 500",
    "INSS": "INSS", "PACIENTE": "Paciente",
}

MEIOS = {"WHATSAPP": "WhatsApp", "LIGACAO": "Ligação", "CALLIX": "Callix", "3C": "3C"}


def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def key(v) -> str:
    """Chave normalizada: sem acento, maiúscula, espaços colapsados."""
    if v is None:
        return ""
    return re.sub(r"\s+", " ", strip_accents(str(v)).strip().upper())


def clean(v):
    if v is None:
        return None
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s or None


def to_date(v) -> str | None:
    if isinstance(v, dt.datetime):
        d = v.date()
    elif isinstance(v, dt.date):
        d = v
    elif isinstance(v, (int, float)) and 30000 < v < 60000:  # serial do Excel
        d = (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    elif isinstance(v, str):
        m = re.match(r"^\s*(\d{1,2})/+(\d{1,2})/+(\d{2,4})\s*$", v)
        if not m:
            return None
        dd, mm, yy = (int(x) for x in m.groups())
        if yy < 100:
            yy += 2000
        try:
            d = dt.date(yy, mm, dd)
        except ValueError:
            return None
    else:
        return None
    if d.year < 2020 or d.year > 2100:
        return None
    return d.isoformat()


def to_hour(v) -> int | None:
    if isinstance(v, dt.time):
        return v.hour
    if isinstance(v, dt.datetime):
        return v.hour
    if isinstance(v, (int, float)) and 0 <= v < 1:  # fração do dia
        return int(v * 24)
    m = re.match(r"^\s*(\d{1,2})", str(v or ""))
    if m and int(m.group(1)) < 24:
        return int(m.group(1))
    return None


def to_money(v) -> float | None:
    if isinstance(v, (int, float)):
        return float(v) if v > 0 else None
    s = str(v or "")
    m = re.search(r"\d[\d.,]*", s)
    if not m:
        return None
    n = m.group(0)
    if "," in n:  # formato BR: 1.500,00
        n = n.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+", n):  # 1.500
        n = n.replace(".", "")
    try:
        val = float(n)
    except ValueError:
        return None
    return val if val > 0 else None


def norm_map(v, table: dict) -> str | None:
    k = key(v)
    if not k:
        return None
    return table.get(k, clean(v).title())


def norm_crc(v) -> str | None:
    s = clean(v)
    if not s or key(s) in {"CONSULTOR (CRC)", "CRC"}:
        return None
    k = key(s)
    if k in {"JAMILY", "JAMYLE"}:
        return "Jamyle"
    return s.title()


def norm_status(v) -> str | None:
    k = key(v)
    if not k:
        return None
    k = re.sub(r"\s*/\s*", "/", k)
    return {
        "COMPARECEU": "Compareceu", "COMPARECIDO": "Compareceu",
        "CONFIRMOU": "Confirmou", "DESMARCOU": "Desmarcou", "FALTOU": "Faltou",
        "REAGENDADO": "Reagendado", "CANCELOU": "Cancelou",
        "CANCELOU/NAO RETORNAR": "Cancelou / não retornar",
    }.get(k, clean(v).capitalize())


def norm_sim_nao(v) -> bool | None:
    k = key(v)
    if not k:
        return None
    if k.startswith("S"):
        return True
    if k.startswith("N"):
        return False
    return None


def norm_fechou(v) -> tuple[str | None, str | None]:
    """Retorna (fechou, motivo). fechou: 'Sim' | 'Não' | None (pendente)."""
    k = key(v)
    if not k:
        return None, None
    if k in {"SIM", "S", "FECHOU"}:
        return "Sim", None
    if k in {"NAO", "N", "NAO FECHOU"}:
        return "Não", None
    # Qualquer outro texto é o motivo de não fechar (ex.: "Mora longe").
    # Curto e categórico; textos longos não são exportados (podem conter dados pessoais).
    motivo = clean(v)
    return "Não", (motivo.capitalize() if motivo and len(motivo) <= 40 else "Outro motivo")


def load_workbook(args) -> openpyxl.Workbook:
    if args.xlsx:
        return openpyxl.load_workbook(args.xlsx, data_only=True, read_only=True)
    file_id = os.environ.get("SHEET_FILE_ID", "").strip()
    if not file_id:
        sys.exit("Defina SHEET_FILE_ID (ID do arquivo no Google Drive) ou use --xlsx.")
    # Planilha Google nativa (export) ou arquivo .xlsx no Drive (download direto).
    urls = [
        f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx",
        f"https://drive.google.com/uc?export=download&id={file_id}",
    ]
    data = b""
    for url in urls:
        req = urllib.request.Request(url, headers={"User-Agent": "dashboard-builder"})
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                data = r.read()
        except Exception:
            continue
        if data.startswith(b"PK"):
            break
    if not data.startswith(b"PK"):
        sys.exit("O download não retornou um .xlsx. Verifique se o arquivo está compartilhado "
                 "como 'Qualquer pessoa com o link'.")
    return openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", help="caminho local da planilha (opcional)")
    ap.add_argument("--out", default="site/data.json")
    args = ap.parse_args()

    wb = load_workbook(args)
    if SHEET not in wb.sheetnames:
        sys.exit(f"Aba '{SHEET}' não encontrada. Abas: {wb.sheetnames}")
    ws = wb[SHEET]

    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        row = list(row) + [None] * (16 - len(row))
        data_agend = to_date(row[COL_DATA_AGEND])
        # Linha válida = tem data do agendamento ou nome do paciente.
        if not data_agend and not clean(row[COL_NOME]):
            continue
        fechou, motivo = norm_fechou(row[COL_FECHOU])
        records.append({
            "a": data_agend,                              # data em que a CRC agendou
            "c": to_date(row[COL_DATA_CONSULTA]),         # data da consulta
            "h": to_hour(row[COL_HORA]),
            "camp": norm_map(row[COL_CAMPANHA], CAMPANHAS),
            "crc": norm_crc(row[COL_CRC]),
            "clin": norm_map(row[COL_CLINICA], CLINICAS),
            "conf": norm_sim_nao(row[COL_CONFIRMACAO]),
            "meio": norm_map(row[COL_MEIO], MEIOS),
            "st": norm_status(row[COL_STATUS]),
            "fechou": fechou,
            "motivo": motivo,
            "orc": bool(clean(row[COL_ORCAMENTO])),
            "orcValor": to_money(row[COL_ORCAMENTO]),
            "venda": to_money(row[COL_VALOR]),
        })

    out = {
        "updatedAt": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "sheet": SHEET,
        "records": records,
    }
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"{len(records)} agendamentos exportados para {args.out}")


if __name__ == "__main__":
    main()
