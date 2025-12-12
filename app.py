import json
from io import BytesIO
from datetime import datetime
from typing import Any, Dict, Tuple

import pandas as pd
import streamlit as st


# ------------------------- Helpers -------------------------

def _safe_dt(dt_str: str | None):
    """Converte ISO datetime do Trello para datetime (ou None)."""
    if not dt_str:
        return None
    try:
        return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
    except Exception:
        return dt_str  # fallback: mantém string


def _label_display(lbl: dict) -> str:
    """Exibe label do Trello com fallback em color caso name esteja vazio."""
    name = (lbl.get("name") or "").strip()
    if name:
        return name
    color = (lbl.get("color") or "").strip()
    return f"(label:{color})" if color else "(label)"


def _ensure_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Garante que um DataFrame não tenha colunas duplicadas.
    (PyArrow/Streamlit quebra quando há colunas repetidas.)
    """
    cols = list(df.columns)
    seen = {}
    new_cols = []
    for c in cols:
        if c not in seen:
            seen[c] = 0
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}__dup{seen[c]}")
    df = df.copy()
    df.columns = new_cols
    return df


# ------------------------- Parser -------------------------

def parse_trello_export(data: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    # Mapas auxiliares
    lists_map = {l.get("id"): l.get("name") for l in data.get("lists", []) or []}
    members_map = {
        m.get("id"): (m.get("fullName") or m.get("username") or m.get("id"))
        for m in (data.get("members", []) or [])
    }
    labels_map = {lb.get("id"): _label_display(lb) for lb in (data.get("labels", []) or [])}

    # -------- Checklists (1 linha por checklist) e itens (1 linha por item)
    checklist_rows = []
    checkitem_rows = []

    for cl in (data.get("checklists", []) or []):
        cl_id = cl.get("id")
        card_id = cl.get("idCard")
        cl_name = cl.get("name")

        checklist_rows.append(
            {
                "checklist_id": cl_id,
                "card_id": card_id,
                "checklist_name": cl_name,
                "checklist_pos": cl.get("pos"),
            }
        )

        for it in (cl.get("checkItems", []) or []):
            checkitem_rows.append(
                {
                    "checklist_id": cl_id,
                    "checklist_name": cl_name,
                    "card_id": card_id,
                    "checkitem_id": it.get("id"),
                    "checkitem_name": it.get("name"),
                    "state": it.get("state"),
                    "checkitem_pos": it.get("pos"),
                    "checkitem_due": _safe_dt(it.get("due")),
                }
            )

    df_checklists = pd.DataFrame(checklist_rows)
    df_checkitems = pd.DataFrame(checkitem_rows)

    # -------- Cards (1 linha por card)
    card_rows = []
    for c in (data.get("cards", []) or []):
        card_id = c.get("id")
        id_list = c.get("idList")

        # Labels no card podem vir como lista de dicts; tenta mapear por id, senão usa fallback
        card_labels = []
        for lb in (c.get("labels", []) or []):
            lb_id = lb.get("id")
            if lb_id and lb_id in labels_map:
                card_labels.append(labels_map[lb_id])
            else:
                card_labels.append(_label_display(lb))
        card_labels = sorted({x for x in card_labels if x})

        card_members = [members_map.get(mid, mid) for mid in (c.get("idMembers", []) or [])]

        card_rows.append(
            {
                "card_id": card_id,
                "idShort": c.get("idShort"),
                "card_name": c.get("name"),
                "list_id": id_list,
                "list_name": lists_map.get(id_list),
                "card_closed": c.get("closed"),
                "card_desc": c.get("desc"),
                "url": c.get("url"),
                "shortLink": c.get("shortLink"),
                "card_due": _safe_dt(c.get("due")),
                "card_start": _safe_dt(c.get("start")),
                "card_dateLastActivity": _safe_dt(c.get("dateLastActivity")),
                "labels": ", ".join(card_labels),
                "members": ", ".join(card_members),
            }
        )

    df_cards = pd.DataFrame(card_rows)

    # -------- Agregado de checklist (texto) por card para a aba Cards
    if not df_checkitems.empty:
        df_ci_sorted = df_checkitems.sort_values(
            by=["card_id", "checklist_name", "checkitem_pos"], na_position="last"
        )
        agg = (
            df_ci_sorted.groupby("card_id", dropna=False)
            .apply(
                lambda g: "\n".join(
                    [f"{r.checklist_name} :: {r.checkitem_name} [{r.state}]"
                     for r in g.itertuples(index=False)]
                )
            )
            .reset_index(name="checklist_items")
        )
    else:
        agg = pd.DataFrame(columns=["card_id", "checklist_items"])

    if not df_cards.empty:
        df_cards = df_cards.merge(agg, how="left", on="card_id")

    # -------- FlatExport (1 linha por item de checklist, com dados do card)
    if not df_checkitems.empty and not df_cards.empty:
        df_flat = df_checkitems.merge(df_cards, how="left", on="card_id")

        preferred = [
            "card_id", "idShort", "card_name", "list_name", "card_closed", "labels", "members",
            "url", "card_dateLastActivity", "card_start", "card_due",
            "checklist_id", "checklist_name", "checklist_pos",
            "checkitem_id", "checkitem_name", "state", "checkitem_pos", "checkitem_due",
            "card_desc",
        ]
        cols = [c for c in preferred if c in df_flat.columns] + [c for c in df_flat.columns if c not in preferred]
        df_flat = df_flat[cols]
    else:
        df_flat = pd.DataFrame(
            columns=[
                "card_id", "idShort", "card_name", "list_name", "card_closed", "labels", "members",
                "url", "card_dateLastActivity", "card_start", "card_due",
                "checklist_id", "checklist_name", "checklist_pos",
                "checkitem_id", "checkitem_name", "state", "checkitem_pos", "checkitem_due",
                "card_desc",
            ]
        )

    # Segurança extra: garante que nenhuma aba fique com colunas duplicadas
    df_cards = _ensure_unique_columns(df_cards)
    df_checklists = _ensure_unique_columns(df_checklists)
    df_checkitems = _ensure_unique_columns(df_checkitems)
    df_flat = _ensure_unique_columns(df_flat)

    return df_cards, df_checkitems, df_checklists, df_flat


def to_excel_bytes(df_cards: pd.DataFrame, df_checkitems: pd.DataFrame, df_checklists: pd.DataFrame, df_flat: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_cards.to_excel(writer, index=False, sheet_name="Cards")
        df_checklists.to_excel(writer, index=False, sheet_name="Checklists")
        df_checkitems.to_excel(writer, index=False, sheet_name="ChecklistItems")
        df_flat.to_excel(writer, index=False, sheet_name="FlatExport")
    return output.getvalue()


# ------------------------- Streamlit UI -------------------------

st.set_page_config(page_title="Trello JSON → Excel", layout="wide")
st.title("Trello JSON → Excel (Cards + Checklists + FlatExport)")

uploaded = st.file_uploader("Envie o JSON exportado do Trello", type=["json"])

if uploaded:
    try:
        data = json.load(uploaded)
    except Exception as e:
        st.error(f"Não consegui ler o JSON: {e}")
        st.stop()

    df_cards, df_checkitems, df_checklists, df_flat = parse_trello_export(data)

    st.subheader("Prévia - FlatExport (1 linha por item de checklist)")
    st.dataframe(df_flat.head(200), use_container_width=True)

    with st.expander("Prévia - Cards"):
        st.dataframe(df_cards.head(200), use_container_width=True)

    with st.expander("Prévia - ChecklistItems"):
        st.dataframe(df_checkitems.head(200), use_container_width=True)

    with st.expander("Prévia - Checklists"):
        st.dataframe(df_checklists.head(200), use_container_width=True)

    xlsx_bytes = to_excel_bytes(df_cards, df_checkitems, df_checklists, df_flat)

    st.download_button(
        label="Baixar Excel (.xlsx)",
        data=xlsx_bytes,
        file_name="trello_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Faça upload do JSON do Trello para gerar o Excel.")
