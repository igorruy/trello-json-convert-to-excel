import json
from io import BytesIO
from datetime import datetime
from typing import Any, Dict, Tuple

import pandas as pd
import streamlit as st


# =========================
# Helpers – Excel
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


def sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    out = out.map(_sanitize_value)
    return out


def export_excel(dfs: Dict[str, pd.DataFrame]) -> bytes:
    engine = _pick_excel_engine()
    output = BytesIO()

    with pd.ExcelWriter(output, engine=engine) as writer:
        for sheet, df in dfs.items():
            df2 = sanitize_df(df)
            # 🔒 garante que a aba exista mesmo vazia
            if df2 is None or df2.empty:
                pd.DataFrame({"info": ["Sem dados"]}).to_excel(
                    writer, sheet_name=sheet, index=False
                )
            else:
                df2.to_excel(writer, sheet_name=sheet, index=False)

    return output.getvalue()


# =========================
# Helpers – Trello
# =========================

def _safe_dt(val):
    if not val:
        return None
    try:
        return datetime.fromisoformat(val.replace("Z", "+00:00")).replace(tzinfo=None)
    except Exception:
        return val


def _label_display(lbl):
    return lbl.get("name") or f"(label:{lbl.get('color')})"


# =========================
# Parser principal
# =========================

def parse_trello(data: dict) -> Tuple[pd.DataFrame, pd.DataFrame]:
    lists = {l["id"]: l["name"] for l in data.get("lists", [])}
    members = {m["id"]: m.get("fullName") or m.get("username") for m in data.get("members", [])}
    labels_map = {l["id"]: _label_display(l) for l in data.get("labels", [])}

    # ---- Cards
    cards = []
    for c in data.get("cards", []):
        card_labels = []
        for lb in c.get("labels", []):
            card_labels.append(labels_map.get(lb.get("id"), _label_display(lb)))

        cards.append({
            "card_id": c["id"],
            "card_name": c["name"],
            "list_name": lists.get(c.get("idList")),
            "labels": ", ".join(sorted(set(card_labels))),
            "members": ", ".join(members.get(mid, mid) for mid in c.get("idMembers", [])),
            "card_due": _safe_dt(c.get("due")),
            "url": c.get("url"),
        })

    df_cards = pd.DataFrame(cards)

    # ---- Checklist items
    items = []
    for cl in data.get("checklists", []):
        for it in cl.get("checkItems", []):
            items.append({
                "card_id": cl.get("idCard"),
                "checklist_name": cl.get("name"),
                "checkitem_name": it.get("name"),
                "state": it.get("state"),
                "checkitem_due": _safe_dt(it.get("due")),
                "responsavel": members.get(it.get("idMember")),
            })

    df_items = pd.DataFrame(items)

    # ---- ABA GERAL (SEM FILTRO)
    if not df_items.empty:
        df_geral = df_items.merge(df_cards, on="card_id", how="left")
    else:
        # garante aba Geral mesmo sem checklist
        df_geral = df_cards.copy()

    # ---- ABA EXPLORE (pendências)
    if not df_geral.empty and "state" in df_geral.columns:
        df_explore = df_geral[df_geral["state"] != "complete"].copy()
    else:
        df_explore = pd.DataFrame(columns=df_geral.columns)

    return df_geral, df_explore


# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title="Trello → Excel (Geral / Explore)", layout="wide")
st.title("Trello JSON → Excel")
st.caption("Abas: Geral (tudo) | Explore (pendências gerenciais)")

uploaded = st.file_uploader("Envie o JSON exportado do Trello", type=["json"])

if uploaded:
    data = json.load(uploaded)

    df_geral, df_explore = parse_trello(data)

    # -------- Filtro Explore
    st.subheader("Explore – Pendências")

    labels = sorted(
        {l.strip() for x in df_explore.get("labels", []) for l in str(x).split(",") if l.strip()}
    )
    selected = st.multiselect("Filtrar por label", labels)

    df_explore_f = df_explore.copy()
    if selected:
        df_explore_f = df_explore_f[
            df_explore_f["labels"].apply(
                lambda x: any(l in str(x) for l in selected)
            )
        ]

    st.dataframe(df_explore_f, use_container_width=True)

    with st.expander("Visualizar aba Geral"):
        st.dataframe(df_geral, use_container_width=True)

    # -------- Exportação
    excel_bytes = export_excel({
        "Geral": df_geral,
        "Explore": df_explore_f,
    })

    st.download_button(
        "Baixar Excel (Geral + Explore)",
        data=excel_bytes,
        file_name="trello_geral_explore.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Faça upload do JSON do Trello.")
