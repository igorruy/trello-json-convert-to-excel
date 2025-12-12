import json
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st


def _safe_dt(dt_str: str | None):
    if not dt_str:
        return None
    try:
        return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
    except Exception:
        return dt_str


def parse_trello_export(data: dict) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    # Mapas auxiliares
    lists_map = {l.get("id"): l.get("name") for l in data.get("lists", [])}
    members_map = {m.get("id"): m.get("fullName") or m.get("username") for m in data.get("members", [])}

    def label_display(lbl: dict) -> str:
        name = (lbl.get("name") or "").strip()
        if name:
            return name
        color = (lbl.get("color") or "").strip()
        return f"(label:{color})" if color else "(label)"

    labels_map = {lb.get("id"): label_display(lb) for lb in data.get("labels", [])}

    # -------- Checklists e itens
    checklist_rows = []
    checkitem_rows = []

    for cl in data.get("checklists", []) or []:
        cl_id = cl.get("id")
        card_id = cl.get("idCard")
        cl_name = cl.get("name")

        checklist_rows.append(
            {
                "checklist_id": cl_id,
                "card_id": card_id,
                "checklist_name": cl_name,
                "pos": cl.get("pos"),
            }
        )

        for it in cl.get("checkItems", []) or []:
            checkitem_rows.append(
                {
                    "checklist_id": cl_id,
                    "checklist_name": cl_name,
                    "card_id": card_id,
                    "checkitem_id": it.get("id"),
                    "checkitem_name": it.get("name"),
                    "state": it.get("state"),
                    "pos": it.get("pos"),
                    "due": _safe_dt(it.get("due")),
                }
            )

    df_checklists = pd.DataFrame(checklist_rows)
    df_checkitems = pd.DataFrame(checkitem_rows)

    # Agregado de itens por card (para coluna em Cards)
    if not df_checkitems.empty:
        df_checkitems_sorted = df_checkitems.sort_values(
            by=["card_id", "checklist_name", "pos"], na_position="last"
        )
        agg = (
            df_checkitems_sorted.groupby("card_id", dropna=False)
            .apply(
                lambda g: "\n".join(
                    [
                        f"{row.checklist_name} :: {row.checkitem_name} [{row.state}]"
                        for row in g.itertuples(index=False)
                    ]
                )
            )
            .reset_index(name="checklist_items")
        )
    else:
        agg = pd.DataFrame(columns=["card_id", "checklist_items"])

    # -------- Cards
    card_rows = []
    for c in data.get("cards", []) or []:
        card_id = c.get("id")
        id_list = c.get("idList")

        card_labels = []
        for lb in c.get("labels", []) or []:
            lb_id = lb.get("id")
            if lb_id and lb_id in labels_map:
                card_labels.append(labels_map[lb_id])
            else:
                card_labels.append(label_display(lb))
        card_labels = sorted({x for x in card_labels if x})

        card_members = [members_map.get(mid, mid) for mid in (c.get("idMembers", []) or [])]

        card_rows.append(
            {
                "card_id": card_id,
                "idShort": c.get("idShort"),
                "card_name": c.get("name"),
                "list_id": id_list,
                "list_name": lists_map.get(id_list),
                "closed": c.get("closed"),
                "desc": c.get("desc"),
                "url": c.get("url"),
                "shortLink": c.get("shortLink"),
                "due": _safe_dt(c.get("due")),
                "start": _safe_dt(c.get("start")),
                "dateLastActivity": _safe_dt(c.get("dateLastActivity")),
                "labels": ", ".join(card_labels),
                "members": ", ".join(card_members),
            }
        )

    df_cards = pd.DataFrame(card_rows)

    if not df_cards.empty:
        df_cards = df_cards.merge(agg, how="left", on="card_id")

    # -------- FlatExport (1 linha por item de checklist, com dados do card)
    if not df_checkitems.empty and not df_cards.empty:
        df_flat = df_checkitems.merge(
            df_cards,
            how="left",
            on="card_id",
            suffixes=("", "_card"),
        )

        # Reordena colunas (opcional, mas melhora leitura)
        preferred = [
            "card_id", "idShort", "card_name", "list_name", "closed", "labels", "members",
            "url", "dateLastActivity", "start", "due",
            "checklist_id", "checklist_name",
            "checkitem_id", "checkitem_name", "state", "pos", "due",
            "desc",
        ]
        cols = [c for c in preferred if c in df_flat.columns] + [c for c in df_flat.columns if c not in preferred]
        df_flat = df_flat[cols]
    else:
        # Se não houver checklists, ainda assim gera aba vazia com colunas esperadas
        df_flat = pd.DataFrame(
            columns=[
                "card_id", "idShort", "card_name", "list_name", "closed", "labels", "members",
                "url", "dateLastActivity", "start", "due",
                "checklist_id", "checklist_name",
                "checkitem_id", "checkitem_name", "state", "pos", "due",
                "desc",
            ]
        )

    return df_cards, df_checkitems, df_checklists, df_flat


def to_excel_bytes(df_cards: pd.DataFrame, df_checkitems: pd.DataFrame, df_checklists: pd.DataFrame, df_flat: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_cards.to_excel(writer, index=False, sheet_name="Cards")
        df_checkitems.to_excel(writer, index=False, sheet_name="ChecklistItems")
        df_checklists.to_excel(writer, index=False, sheet_name="Checklists")
        df_flat.to_excel(writer, index=False, sheet_name="FlatExport")
    return output.getvalue()


# ---------------- STREAMLIT UI ----------------
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
    st.dataframe(df_flat.head(100), use_container_width=True)

    xlsx_bytes = to_excel_bytes(df_cards, df_checkitems, df_checklists, df_flat)

    st.download_button(
        label="Baixar Excel (.xlsx)",
        data=xlsx_bytes,
        file_name="trello_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
