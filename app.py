import json
from io import BytesIO
from datetime import datetime
from typing import Any, Dict

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
        return v.replace(tzinfo=None) if v.tzinfo else v
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
    if df is None or df.empty:
        return df
    df = df.copy().map(_sanitize_value)
    return _ensure_unique_columns(df)


def dfs_to_xlsx_bytes(dfs: Dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine=_pick_excel_engine()) as writer:
        for name, df in dfs.items():
            df2 = sanitize_df(df)
            if df2 is None or df2.empty:
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


def _is_missing_date(x) -> bool:
    # cobre None, NaT, NaN
    return x is None or (isinstance(x, float) and pd.isna(x)) or pd.isna(x)


def _calc_prazo(card_due, dueComplete, report_dt: datetime):
    """
    Regra:
      - card fechado => Concluído
      - sem due (None/NaT) => Em dia
      - report_dt > due => Em atraso
      - senão => Em dia
    """
    if bool(dueComplete):
        return "Concluído"

    if _is_missing_date(card_due):
        return "Em dia"

    # card_due pode ser datetime ou Timestamp
    try:
        due_date = card_due.date()  # datetime
    except Exception:
        # Timestamp/valores estranhos
        due_date = pd.to_datetime(card_due, errors="coerce")
        if pd.isna(due_date):
            return "Em dia"
        due_date = due_date.date()

    return "Em atraso" if report_dt.date() > due_date else "Em dia"


# =========================
# Parser principal
# =========================

def parse_trello(data: dict, report_dt: datetime):
    lists = {l["id"]: l["name"] for l in data.get("lists", []) or []}
    members = {m["id"]: (m.get("fullName") or m.get("username") or m["id"]) for m in data.get("members", []) or []}
    labels = {l["id"]: _label_display(l) for l in data.get("labels", []) or []}

    # Cards
    cards = []
    for c in data.get("cards", []) or []:
        card_labels = []
        for lb in c.get("labels", []) or []:
            lb_id = lb.get("id")
            card_labels.append(labels.get(lb_id, _label_display(lb)))

        id_members = c.get("idMembers", []) or []
        cards.append({
            "card_id": c.get("id"),
            "card_name": c.get("name"),
            "list_name": lists.get(c.get("idList")),
            "dueComplete": bool(c.get("closed")),
            "card_due": _safe_dt(c.get("due")),
            "labels": ", ".join(sorted(set([x for x in card_labels if x]))),
            "members": ", ".join([members.get(m, m) for m in id_members]),
            "url": c.get("url"),
        })

    df_cards = pd.DataFrame(cards)

    # Prazo (sem apply vulnerável a NaT)
    if not df_cards.empty:
        df_cards["Prazo"] = [
            _calc_prazo(due, closed, report_dt)
            for due, closed in zip(df_cards.get("card_due", []), df_cards.get("dueComplete", []))
        ]

    # Checklist Items (simples)
    items = []
    for cl in data.get("checklists", []) or []:
        for it in cl.get("checkItems", []) or []:
            items.append({
                "card_id": cl.get("idCard"),
                "checklist_name": cl.get("name"),
                "checkitem_name": it.get("name"),
                "state": it.get("state"),
                "responsavel": members.get(it.get("idMember")),
            })

    df_items = pd.DataFrame(items)

    # FlatExport SEM filtro (todos cards + todos itens)
    if not df_cards.empty:
        if not df_items.empty:
            df_flat = df_cards.merge(df_items, on="card_id", how="left")
        else:
            df_flat = df_cards.copy()
    else:
        df_flat = pd.DataFrame()

    # Explore (pendentes apenas)
    if not df_flat.empty and "state" in df_flat.columns:
        df_explore = df_flat[df_flat["state"].fillna("") != "complete"].copy()
    else:
        df_explore = pd.DataFrame()

    return df_cards, df_items, df_flat, df_explore


# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title="Trello → Excel (sem limite)", layout="wide")
st.title("Trello JSON → Excel – SEM LIMITE DE LINHAS")
st.caption("Corrigido para lidar com card_due vazio/NaT. FlatExport sem filtro.")

uploaded = st.file_uploader("Envie o JSON do Trello", type=["json"])

if uploaded:
    data = json.load(uploaded)
    report_dt = datetime.now()

    df_cards, df_items, df_flat, df_explore = parse_trello(data, report_dt)

    # Debug (ajuda a validar que não há corte)
    st.subheader("Contagens")
    st.metric("Cards no JSON", len(data.get("cards", []) or []))
    st.metric("Cards (df_cards)", len(df_cards))
    st.metric("Linhas FlatExport", len(df_flat))
    st.metric("Linhas Explore", len(df_explore))

    st.subheader("FlatExport")
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
