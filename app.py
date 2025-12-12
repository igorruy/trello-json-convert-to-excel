import json
from io import BytesIO
from datetime import datetime
from typing import Any, Dict, Tuple

import pandas as pd
import streamlit as st


# =========================
# Excel / Sanitização
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
        return v.tz_convert(None).to_pydatetime() if v.tz else v.to_pydatetime()
    if isinstance(v, datetime):
        return v.replace(tzinfo=None)
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    return v


def _ensure_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = []
    seen = {}
    for c in df.columns:
        if c not in seen:
            seen[c] = 0
            cols.append(c)
        else:
            seen[c] += 1
            cols.append(f"{c}__dup{seen[c]}")
    df = df.copy()
    df.columns = cols
    return df


def sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.map(_sanitize_value)
    return _ensure_unique_columns(df)


def dfs_to_xlsx_bytes(dfs: Dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine=_pick_excel_engine()) as writer:
        for name, df in dfs.items():
            df2 = sanitize_df(df)
            if df2.empty:
                pd.DataFrame({"info": ["Sem dados"]}).to_excel(writer, sheet_name=name[:31], index=False)
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


def _calc_prazo(card_due, card_closed, report_dt):
    if card_closed:
        return "Concluído"
    if not card_due:
        return "Em dia"
    return "Em atraso" if report_dt.date() > card_due.date() else "Em dia"


# =========================
# Parser principal
# =========================

def parse_trello(data: dict, report_dt: datetime):
    lists = {l["id"]: l["name"] for l in data.get("lists", [])}
    members = {m["id"]: m.get("fullName") or m.get("username") for m in data.get("members", [])}
    labels = {l["id"]: _label_display(l) for l in data.get("labels", [])}

    # Cards
    cards = []
    for c in data.get("cards", []):
        card_labels = [_label_display(lb) for lb in c.get("labels", [])]
        cards.append({
            "card_id": c["id"],
            "card_name": c["name"],
            "list_name": lists.get(c.get("idList")),
            "card_closed": c.get("closed"),
            "card_due": _safe_dt(c.get("due")),
            "labels": ", ".join(sorted(set(card_labels))),
            "members": ", ".join(members.get(m) for m in c.get("idMembers", [])),
            "url": c.get("url"),
        })

    df_cards = pd.DataFrame(cards)
    df_cards["Prazo"] = df_cards.apply(
        lambda r: _calc_prazo(r["card_due"], r["card_closed"], report_dt), axis=1
    )

    # Checklists / Items
    items = []
    for cl in data.get("checklists", []):
        for it in cl.get("checkItems", []):
            items.append({
                "card_id": cl["idCard"],
                "checklist_name": cl["name"],
                "checkitem_name": it["name"],
                "state": it["state"],
                "responsavel": members.get(it.get("idMember")),
            })

    df_items = pd.DataFrame(items)

    # FlatExport SEM FILTRO
    df_flat = df_cards.merge(df_items, on="card_id", how="left")

    # Explore (somente pendentes)
    df_explore = df_flat[df_flat["state"] != "complete"].copy()

    return df_cards, df_items, df_flat, df_explore


# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title="Trello → Excel (sem limite)", layout="wide")
st.title("Trello JSON → Excel – SEM LIMITE DE LINHAS")

uploaded = st.file_uploader("Envie o JSON do Trello", type=["json"])

if uploaded:
    data = json.load(uploaded)
    report_dt = datetime.now()

    df_cards, df_items, df_flat, df_explore = parse_trello(data, report_dt)

    st.subheader("FlatExport (SEM filtro e SEM limite)")
    st.dataframe(df_flat, use_container_width=True, height=700)

    st.subheader("Explore (pendências)")
    st.dataframe(df_explore, use_container_width=True, height=700)

    with st.expander("Cards"):
        st.dataframe(df_cards, use_container_width=True, height=700)

    with st.expander("Checklist Items"):
        st.dataframe(df_items, use_container_width=True, height=700)

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
        "trello_export_completo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Faça upload do JSON do Trello.")
