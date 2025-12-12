import json
from io import BytesIO
from datetime import datetime
from typing import Any, Dict, Tuple

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
    raise ModuleNotFoundError("Instale XlsxWriter ou openpyxl no requirements.txt")


def _sanitize_value(v: Any) -> Any:
    if isinstance(v, pd.Timestamp):
        # remove timezone e converte para python datetime quando possível
        if v.tz is not None:
            return v.tz_convert(None).to_pydatetime()
        return v.to_pydatetime()
    if isinstance(v, datetime):
        return v.replace(tzinfo=None) if v.tzinfo else v
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    return v


def sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    return df.copy().map(_sanitize_value)


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
    return (lbl.get("name") or "").strip() or f"(label:{lbl.get('color')})"


def _is_due_complete(card: dict) -> bool:
    # Correto: somente True explícito
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

def parse_trello(data: dict, report_dt: datetime) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    lists = {l.get("id"): l.get("name") for l in data.get("lists", []) or []}
    members = {
        m.get("id"): (m.get("fullName") or m.get("username") or m.get("id"))
        for m in data.get("members", []) or []
    }
    labels_map = {l.get("id"): _label_display(l) for l in data.get("labels", []) or []}

    # ----------------- Cards
    cards = []
    for c in data.get("cards", []) or []:
        card_labels = []
        for lb in c.get("labels", []) or []:
            lb_id = lb.get("id")
            card_labels.append(labels_map.get(lb_id, _label_display(lb)))

        due_complete = _is_due_complete(c)
        due_dt = _safe_dt(c.get("due"))

        cards.append({
            "card_id": c.get("id"),
            "idShort": c.get("idShort"),
            "card_name": c.get("name"),
            "list_name": lists.get(c.get("idList")),
            "dueComplete": due_complete,
            "card_due": due_dt,
            "Prazo": _calc_prazo(due_dt, due_complete, report_dt),
            "labels": ", ".join(sorted(set([x for x in card_labels if x]))),
            "members": ", ".join([members.get(mid, mid) for mid in (c.get("idMembers", []) or [])]),
            "url": c.get("url"),
            "shortLink": c.get("shortLink"),
            "card_dateLastActivity": _safe_dt(c.get("dateLastActivity")),
            "card_start": _safe_dt(c.get("start")),
            "card_desc": c.get("desc"),
        })

    df_cards = pd.DataFrame(cards)

    # ----------------- Checklists (1 linha por checklist) + resumo
    checklist_rows = []
    for cl in data.get("checklists", []) or []:
        items = cl.get("checkItems", []) or []
        total = len(items)
        concluidos = sum(1 for it in items if (it.get("state") or "").lower() == "complete")
        pendentes = total - concluidos

        checklist_rows.append({
            "checklist_id": cl.get("id"),
            "card_id": cl.get("idCard"),
            "checklist_name": cl.get("name"),
            "checklist_pos": cl.get("pos"),
            "total_itens": total,
            "concluidos": concluidos,
            "pendentes": pendentes,
        })

    df_checklists = pd.DataFrame(checklist_rows)

    # ----------------- ChecklistItems (1 linha por item)
    item_rows = []
    for cl in data.get("checklists", []) or []:
        for it in cl.get("checkItems", []) or []:
            item_rows.append({
                "card_id": cl.get("idCard"),
                "checklist_id": cl.get("id"),
                "checklist_name": cl.get("name"),
                "checkitem_id": it.get("id"),
                "checkitem_name": it.get("name"),
                "state": it.get("state"),
                "checkitem_pos": it.get("pos"),
                "responsavel_id": it.get("idMember"),
                "responsavel": members.get(it.get("idMember")),
            })

    df_items = pd.DataFrame(item_rows)

    # ----------------- FlatExport (SEM filtro): todos os cards + itens (quando existirem)
    if not df_cards.empty and not df_items.empty:
        df_flat = df_cards.merge(df_items, on="card_id", how="left")
    else:
        df_flat = df_cards.copy()

    # ----------------- Explore (pendências): itens incompletos E card não concluído
    if not df_flat.empty and "state" in df_flat.columns:
        df_explore = df_flat[
            (df_flat["state"].fillna("").str.lower() != "complete") &
            (df_flat["dueComplete"] != True)
        ].copy()
    else:
        df_explore = pd.DataFrame()

    return df_cards, df_checklists, df_items, df_flat, df_explore


# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title="Trello → Excel (dueComplete)", layout="wide")
st.title("Trello JSON → Excel")
st.caption("Conclusão baseada em dueComplete (não confunde com arquivamento). Visualizações completas no Streamlit.")

uploaded = st.file_uploader("Envie o JSON do Trello", type=["json"])

if uploaded:
    data = json.load(uploaded)
    report_dt = datetime.now()

    df_cards, df_checklists, df_items, df_flat, df_explore = parse_trello(data, report_dt)

    # ----------------- Filtro por label (Explore)
    st.subheader("Explore (pendências) — filtro por label")

    all_labels = []
    if not df_explore.empty and "labels" in df_explore.columns:
        uniq = set()
        for s in df_explore["labels"].fillna("").astype(str).tolist():
            for part in [p.strip() for p in s.split(",") if p.strip()]:
                uniq.add(part)
        all_labels = sorted(uniq)

    selected_labels = st.multiselect("Filtrar por label", options=all_labels, default=[])

    df_explore_filtered = df_explore.copy()
    if selected_labels and not df_explore_filtered.empty and "labels" in df_explore_filtered.columns:
        def has_any_label(label_str: str) -> bool:
            ls = {p.strip() for p in str(label_str).split(",") if p.strip()}
            return any(l in ls for l in selected_labels)
        df_explore_filtered = df_explore_filtered[df_explore_filtered["labels"].apply(has_any_label)].copy()

    # ----------------- Contagens (diagnóstico rápido)
    st.subheader("Contagens")
    st.metric("Cards no JSON", len(data.get("cards", []) or []))
    st.metric("Cards (df_cards)", len(df_cards))
    st.metric("Cards concluídos (dueComplete=True)", int(df_cards["dueComplete"].sum()) if not df_cards.empty else 0)
    st.metric("Linhas FlatExport", len(df_flat))
    st.metric("Linhas Explore (após filtro)", len(df_explore_filtered))

    # ----------------- Visões no Streamlit (SEM limite)
    st.subheader("FlatExport (sem filtro)")
    st.dataframe(df_flat, use_container_width=True, height=700)

    st.subheader("Explore (pendências)")
    st.dataframe(df_explore_filtered, use_container_width=True, height=700)

    with st.expander("Cards"):
        st.dataframe(df_cards, use_container_width=True, height=700)

    with st.expander("Checklists"):
        st.dataframe(df_checklists, use_container_width=True, height=700)

    with st.expander("ChecklistItems"):
        st.dataframe(df_items, use_container_width=True, height=700)

    # ----------------- Exportação
    st.divider()
    excel_bytes = dfs_to_xlsx_bytes({
        "Cards": df_cards,
        "Checklists": df_checklists,
        "ChecklistItems": df_items,
        "FlatExport": df_flat,
        "Explore": df_explore_filtered,  # exporta com filtro aplicado
    })

    st.download_button(
        "Baixar Excel completo",
        excel_bytes,
        "trello_export_dueComplete.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Faça upload do JSON do Trello.")
