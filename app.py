import json
from io import BytesIO
from datetime import datetime
from typing import Any, Dict

import pandas as pd
import streamlit as st


# =========================
# Excel helpers
# =========================

def _pick_excel_engine() -> str:
    try:
        import xlsxwriter  # noqa
        return "xlsxwriter"
    except Exception:
        pass
    try:
        import openpyxl  # noqa
        return "openpyxl"
    except Exception:
        pass
    raise ModuleNotFoundError("Instale XlsxWriter ou openpyxl")


def _sanitize_value(v: Any) -> Any:
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if isinstance(v, datetime):
        return v.replace(tzinfo=None)
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    return v


def sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    return df.map(_sanitize_value)


def dfs_to_xlsx_bytes(dfs: Dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine=_pick_excel_engine()) as writer:
        for name, df in dfs.items():
            df2 = sanitize_df(df)
            if df2.empty:
                pd.DataFrame({"info": ["Sem dados"]}).to_excel(
                    writer, sheet_name=name[:31], index=False
                )
            else:
                df2.to_excel(writer, sheet_name=name[:31], index=False)
    return output.getvalue()


# =========================
# Trello helpers
# =========================

def _safe_dt(v):
    if not v:
        return None
    try:
        return datetime.fromisoformat(str(v).replace("Z", "+00:00")).replace(tzinfo=None)
    except Exception:
        return None


def _label_display(lbl):
    return lbl.get("name") or f"(label:{lbl.get('color')})"


def _is_due_complete(card: dict) -> bool:
    # REGRA CORRETA: somente True explícito
    return card.get("dueComplete") is True


def _calc_prazo(card_due, due_complete: bool, report_dt: datetime) -> str:
    if due_complete:
        return "Concluído"

    if card_due is None or pd.isna(card_due):
        return "Em dia"

    try:
        return "Em atraso" if report_dt.date() > card_due.date() else "Em dia"
    except Exception:
        return "Em dia"


# =========================
# Parser principal
# =========================

def parse_trello(data: dict, report_dt: datetime):
    lists = {l["id"]: l["name"] for l in data.get("lists", [])}
    members = {
        m["id"]: (m.get("fullName") or m.get("username") or m["id"])
        for m in data.get("members", [])
    }

    # Cards
    cards = []
    for c in data.get("cards", []):
        labels = [_label_display(lb) for lb in c.get("labels", [])]

        due_complete = _is_due_complete(c)
        due_dt = _safe_dt(c.get("due"))

        cards.append({
            "card_id": c.get("id"),
            "card_name": c.get("name"),
            "list_name": lists.get(c.get("idList")),
            "dueComplete": due_complete,
            "card_due": due_dt,
            "Prazo": _calc_prazo(due_dt, due_complete, report_dt),
            "labels": ", ".join(sorted(set(labels))),
            "members": ", ".join(
                members.get(m, m) for m in c.get("idMembers", [])
            ),
            "url": c.get("url"),
        })

    df_cards = pd.DataFrame(cards)

    # Checklist items
    items = []
    for cl in data.get("checklists", []):
        for it in cl.get("checkItems", []):
            items.append({
                "card_id": cl.get("idCard"),
                "checklist_name": cl.get("name"),
                "checkitem_name": it.get("name"),
                "state": it.get("state"),
                "responsavel": members.get(it.get("idMember")),
            })

    df_items = pd.DataFrame(items)

    # Flat (SEM filtro)
    if not df_items.empty:
        df_flat = df_cards.merge(df_items, on="card_id", how="left")
    else:
        df_flat = df_cards.copy()

    # Explore (pendências)
    df_explore = df_flat[
        (df_flat["state"] != "complete") & (df_flat["dueComplete"] != True)
    ].copy()

    return df_cards, df_items, df_flat, df_explore


# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title="Trello → Excel (dueComplete)", layout="wide")
st.title("Trello JSON → Excel")
st.caption("Status de conclusão baseado em dueComplete (corrigido).")

uploaded = st.file_uploader("Envie o JSON do Trello", type=["json"])

if uploaded:
    data = json.load(uploaded)
    report_dt = datetime.now()

    df_cards, df_items, df_flat, df_explore = parse_trello(data, report_dt)

    st.subheader("Resumo")
    st.metric("Cards no JSON", len(data.get("cards", [])))
    st.metric("Cards concluídos (dueComplete)", df_cards["dueComplete"].sum())
    st.metric("Linhas FlatExport", len(df_flat))

    st.subheader("FlatExport (sem filtro)")
    st.dataframe(df_flat, use_container_width=True, height=700)

    st.subheader("Explore (pendências)")
    st.dataframe(df_explore, use_container_width=True, height=700)

    st.divider()

    excel_bytes = dfs_to_xlsx_bytes({
        "Cards": df_cards,
        "ChecklistItems": df_items,
        "FlatExport": df_flat,
        "Explore": df_explore,
    })

    st.download_button(
        "Baixar Excel completo",
        excel_bytes,
        "trello_export_dueComplete.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Faça upload do JSON do Trello.")
