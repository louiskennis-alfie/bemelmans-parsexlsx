from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
import io
from typing import List, Dict, Any

app = FastAPI(title="Excel BOQ Parser")

# Autoriser les appels depuis n8n / front
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # à restreindre en prod
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def extract_visible_rows_from_active_sheet(
    content: bytes,
    max_rows: int | None = None,
) -> Dict[str, Any]:
    try:
        wb = load_workbook(io.BytesIO(content), data_only=True)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Impossible de lire le fichier Excel: {e}")

    ws = wb.active  # dernière feuille ouverte
    sheet_name = ws.title

    rows_out: List[Dict[str, Any]] = []

    for row in ws.iter_rows():
        first_cell = row[0]
        row_idx = first_cell.row

        row_dim = ws.row_dimensions.get(row_idx)
        if row_dim is not None and row_dim.hidden:
            continue

        cells = []
        all_empty = True

        for cell in row:
            col_index = cell.column
            col_letter = get_column_letter(col_index)
            address = f"{col_letter}{cell.row}"

            if isinstance(cell, MergedCell):
                value = None
            else:
                value = cell.value

            if value is not None and str(value).strip() != "":
                all_empty = False

            cells.append(
                {
                    "address": address,
                    "col_index": col_index,
                    "col_letter": col_letter,
                    "row_index": cell.row,
                    "value": value,
                }
            )

        if all_empty:
            continue

        rows_out.append(
            {
                "row_index": row_idx,
                "cells": cells,
            }
        )

        if max_rows is not None and len(rows_out) >= max_rows:
            break

    return {
        "sheet_name": sheet_name,
        "row_count": len(rows_out),
        "rows": rows_out,
    }


@app.post("/parse_excel")
async def parse_excel(
    file: UploadFile = File(...),
    max_rows: int | None = None,
):
    """
    Endpoint principal :
    - reçoit un fichier Excel (champ 'file')
    - lit la dernière feuille ouverte
    - retourne les lignes visibles en JSON
    """
    if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        raise HTTPException(status_code=400, detail="Le fichier doit être un Excel .xlsx / .xlsm / .xltx / .xltm")

    content = await file.read()

    data = extract_visible_rows_from_active_sheet(content, max_rows=max_rows)

    return {
        "filename": file.filename,
        "sheet_name": data["sheet_name"],
        "row_count": data["row_count"],
        "rows": data["rows"],
    }


@app.get("/health")
def health():
    return {"status": "ok"}
