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

def parse_trello(data: dict, report_dt: datetime, derive_cfg: Dict[str, Any] | None = None) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    lists = {l.get("id"): l.get("name") for l in data.get("lists", []) or []}
    members = {
        m.get("id"): (m.get("fullName") or m.get("username") or m.get("id"))
        for m in data.get("members", []) or []
    }
    labels_map = {l.get("id"): _label_display(l) for l in data.get("labels", []) or []}
    board_id = data.get("id")

    custom_fields = data.get("customFields", []) or []
    custom_fields_by_id = {cf.get("id"): cf for cf in custom_fields if cf.get("id")}

    def _custom_field_value(item: dict):
        cf_def = custom_fields_by_id.get(item.get("idCustomField"))
        if not cf_def:
            return None
        value = item.get("value") or {}
        cf_type = cf_def.get("type")
        if cf_type == "text":
            return value.get("text")
        if cf_type == "number":
            return value.get("number")
        if cf_type == "date":
            return _safe_dt(value.get("date"))
        if cf_type == "checkbox":
            return value.get("checked")
        if cf_type == "list":
            option_id = item.get("idValue")
            if not option_id:
                return None
            for opt in cf_def.get("options", []) or []:
                if opt.get("id") == option_id:
                    opt_value = opt.get("value") or {}
                    return opt_value.get("text") or opt_value.get("number")
            return None
        return None

    # ----------------- Cards
    cards = []
    for c in data.get("cards", []) or []:
        card_labels = []
        for lb in c.get("labels", []) or []:
            lb_id = lb.get("id")
            card_labels.append(labels_map.get(lb_id, _label_display(lb)))

        custom_values = {}
        for item in c.get("customFieldItems", []) or []:
            cf_def = custom_fields_by_id.get(item.get("idCustomField"))
            if not cf_def:
                continue
            name = cf_def.get("name")
            if not name:
                continue
            custom_values[name] = _custom_field_value(item)

        due_complete = _is_due_complete(c)
        due_dt = _safe_dt(c.get("due"))

        card_data = {
            "board_id": board_id,
            "card_id": c.get("id"),
            "idShort": c.get("idShort"),
            "card_name": c.get("name"),
            "list_id": c.get("idList"),
            "list_name": lists.get(c.get("idList")),
            "dueComplete": due_complete,
            "card_due": due_dt,
            "Prazo": _calc_prazo(due_dt, due_complete, report_dt),
            "labels": ", ".join(sorted(set([x for x in card_labels if x]))),
            "members": ", ".join([members.get(mid, mid) for mid in (c.get("idMembers", []) or [])]),
            "member_ids": ", ".join([(mid or "") for mid in (c.get("idMembers", []) or [])]),
            "url": c.get("url"),
            "shortLink": c.get("shortLink"),
            "card_dateLastActivity": _safe_dt(c.get("dateLastActivity")),
            "card_start": _safe_dt(c.get("start")),
            "card_desc": c.get("desc"),
        }

        if custom_values:
            card_data.update(custom_values)

        cards.append(card_data)

    df_cards = pd.DataFrame(cards)
    derive_cfg = derive_cfg or {}
    if not df_cards.empty:
        if derive_cfg.get("dias_em_atraso"):
            def _dias_em_atraso(row):
                d = row.get("card_due")
                dc = row.get("dueComplete")
                if d is None or pd.isna(d) or dc is True:
                    return 0
                try:
                    delta = (report_dt.date() - d.date()).days
                    return max(delta, 0)
                except Exception:
                    return 0
            df_cards["dias_em_atraso"] = df_cards.apply(_dias_em_atraso, axis=1)
        if derive_cfg.get("tempo_em_lista"):
            def _tempo_em_lista(row):
                base = row.get("card_start") or row.get("card_dateLastActivity")
                if base is None or pd.isna(base):
                    return None
                try:
                    delta = (report_dt.date() - base.date()).days
                    return max(delta, 0)
                except Exception:
                    return None
            df_cards["tempo_em_lista"] = df_cards.apply(_tempo_em_lista, axis=1)

    # ----------------- Checklists (1 linha por checklist) + resumo
    checklist_rows = []
    card_to_list = {c.get("card_id"): c.get("list_id") for c in cards}
    for cl in data.get("checklists", []) or []:
        items = cl.get("checkItems", []) or []
        total = len(items)
        concluidos = sum(1 for it in items if (it.get("state") or "").lower() == "complete")
        pendentes = total - concluidos

        checklist_rows.append({
            "board_id": board_id,
            "checklist_id": cl.get("id"),
            "card_id": cl.get("idCard"),
            "list_id": card_to_list.get(cl.get("idCard")),
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
                "board_id": board_id,
                "card_id": cl.get("idCard"),
                "list_id": card_to_list.get(cl.get("idCard")),
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

    if not df_flat.empty:
        if {"list_id_x", "list_id_y"}.issubset(df_flat.columns):
            df_flat["list_id"] = df_flat["list_id_x"].fillna(df_flat["list_id_y"])
            df_flat = df_flat.drop(columns=["list_id_x", "list_id_y"])
        if {"board_id_x", "board_id_y"}.issubset(df_flat.columns):
            df_flat["board_id"] = df_flat["board_id_x"].fillna(df_flat["board_id_y"])
            df_flat = df_flat.drop(columns=["board_id_x", "board_id_y"])
        rename_map = {
            "board_id": "Board ID",
            "card_id": "Card ID",
            "idShort": "Número",
            "card_name": "Título",
            "list_id": "Lista ID",
            "list_name": "Lista",
            "dueComplete": "Concluído",
            "card_due": "Data de Entrega",
            "Prazo": "Prazo",
            "labels": "Etiquetas",
            "members": "Membros",
            "member_ids": "IDs de Membros",
            "dias_em_atraso": "Dias em atraso",
            "tempo_em_lista": "Tempo em lista",
            "url": "URL",
            "shortLink": "Link curto",
            "card_dateLastActivity": "Última atividade",
            "card_start": "Início",
            "card_desc": "Descrição",
            "checklist_id": "Checklist ID",
            "checklist_name": "Checklist",
            "checkitem_id": "Item ID",
            "checkitem_name": "Item",
            "state": "Status",
            "checkitem_pos": "Posição do Item",
            "responsavel_id": "Responsável ID",
            "responsavel": "Responsável",
        }
        df_flat = df_flat.rename(columns={k: v for k, v in rename_map.items() if k in df_flat.columns})

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

    st.subheader("Opções de colunas derivadas")
    opt_dias = st.checkbox("dias_em_atraso", value=True)
    opt_tempo = st.checkbox("tempo_em_lista", value=True)
    derive_cfg = {"dias_em_atraso": opt_dias, "tempo_em_lista": opt_tempo}

    df_cards, df_checklists, df_items, df_flat, df_explore = parse_trello(data, report_dt, derive_cfg)

    # ----------------- Contagens (diagnóstico rápido)
    st.subheader("Contagens")
    st.metric("Cards no JSON", len(data.get("cards", []) or []))
    st.metric("Cards (df_cards)", len(df_cards))
    st.metric("Cards concluídos (dueComplete=True)", int(df_cards["dueComplete"].sum()) if not df_cards.empty else 0)
    st.metric("Linhas FlatExport", len(df_flat))

    # ----------------- Visões no Streamlit (SEM limite)
    st.subheader("FlatExport (sem filtro)")
    st.dataframe(df_flat, use_container_width=True, height=700)

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
        "FlatExport": df_flat
    })

    st.download_button(
        "Baixar Excel completo",
        excel_bytes,
        "trello_export_dueComplete.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Faça upload do JSON do Trello.")
