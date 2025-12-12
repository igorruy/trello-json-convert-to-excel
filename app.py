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
        import xlsxwriter  # noqa: F401
        return "xlsxwriter"
    except Exception:
        pass

    try:
        import openpyxl  # noqa: F401
        return "openpyxl"
    except Exception:
        pass

    raise ModuleNotFoundError(
        "Nenhum engine de Excel disponível. Instale 'XlsxWriter' (recomendado) ou 'openpyxl' no requirements.txt."
    )


def _sanitize_value_for_excel(v: Any) -> Any:
    if isinstance(v, pd.Timestamp):
        if v.tz is not None:
            return v.tz_convert(None).to_pydatetime()
        return v.to_pydatetime()

    if isinstance(v, datetime):
        if v.tzinfo is not None:
            return v.replace(tzinfo=None)
        return v

    if isinstance(v, (dict, list)):
        try:
            return json.dumps(v, ensure_ascii=False)
        except Exception:
            return str(v)

    return v


def _ensure_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
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
    out = df.copy()
    out.columns = new_cols
    return out


def sanitize_df_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    out = out.map(_sanitize_value_for_excel)
    out = _ensure_unique_columns(out)
    return out


def dfs_to_xlsx_bytes(dfs: Dict[str, pd.DataFrame]) -> bytes:
    engine = _pick_excel_engine()
    output = BytesIO()
    with pd.ExcelWriter(output, engine=engine) as writer:
        for sheet, df in dfs.items():
            df2 = sanitize_df_for_excel(df)
            if df2 is None or df2.empty:
                pd.DataFrame({"info": ["Sem dados"]}).to_excel(writer, index=False, sheet_name=sheet[:31])
            else:
                df2.to_excel(writer, index=False, sheet_name=sheet[:31])
    return output.getvalue()


def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    return dfs_to_xlsx_bytes({sheet_name: df})


# =========================
# Trello parsing / Prazo
# =========================

def _safe_dt(dt_str: str | None):
    if not dt_str:
        return None
    try:
        dt = datetime.fromisoformat(str(dt_str).replace("Z", "+00:00"))
        return dt.replace(tzinfo=None) if dt.tzinfo else dt
    except Exception:
        return None


def _label_display(lbl: dict) -> str:
    name = (lbl.get("name") or "").strip()
    if name:
        return name
    color = (lbl.get("color") or "").strip()
    return f"(label:{color})" if color else "(label)"


def _calc_prazo(card_due: datetime | None, card_closed: bool | None, report_dt: datetime) -> str:
    """
    Nova regra:
      - Se card_closed == True -> "Concluído"
      - Se não houver card_due -> "Em dia"
      - Se report_dt.date() > card_due.date() -> "Em atraso"
      - Senão -> "Em dia"
    """
    if bool(card_closed):
        return "Concluído"
    if not card_due:
        return "Em dia"
    return "Em atraso" if report_dt.date() > card_due.date() else "Em dia"


def parse_trello_export(
    data: dict,
    report_dt: datetime
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:

    lists_map = {l.get("id"): l.get("name") for l in data.get("lists", []) or []}
    members_map = {
        m.get("id"): (m.get("fullName") or m.get("username") or m.get("id"))
        for m in data.get("members", []) or []
    }
    labels_map = {lb.get("id"): _label_display(lb) for lb in data.get("labels", []) or []}

    # ----------------- Checklists e itens
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
                "checklist_pos": cl.get("pos"),
            }
        )

        for it in cl.get("checkItems", []) or []:
            id_member = it.get("idMember")
            responsavel = members_map.get(id_member) if id_member else None

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
                    "responsavel_id": id_member,
                    "responsavel": responsavel,
                }
            )

    df_checklists = pd.DataFrame(checklist_rows)
    df_checkitems = pd.DataFrame(checkitem_rows)

    # ----------------- Cards
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
                card_labels.append(_label_display(lb))
        card_labels = sorted({x for x in card_labels if x})

        card_members = [members_map.get(mid, mid) for mid in (c.get("idMembers", []) or [])]

        card_due = _safe_dt(c.get("due"))
        card_closed = bool(c.get("closed"))

        card_rows.append(
            {
                "card_id": card_id,
                "idShort": c.get("idShort"),
                "card_name": c.get("name"),
                "list_id": id_list,
                "list_name": lists_map.get(id_list),
                "card_closed": card_closed,
                "card_desc": c.get("desc"),
                "url": c.get("url"),
                "shortLink": c.get("shortLink"),
                "card_due": card_due,
                "card_start": _safe_dt(c.get("start")),
                "card_dateLastActivity": _safe_dt(c.get("dateLastActivity")),
                "labels": ", ".join(card_labels),
                "members": ", ".join(card_members),
                "Prazo": _calc_prazo(card_due, card_closed, report_dt),
            }
        )

    df_cards = pd.DataFrame(card_rows)

    # ----------------- Agregado checklist_items em Cards
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

    # ----------------- FlatExport
    if not df_checkitems.empty and not df_cards.empty:
        df_flat = df_checkitems.merge(df_cards, how="left", on="card_id")

        # Prazo agora vem do card (card_closed -> Concluído, senão atraso/dia)
        if "Prazo" not in df_flat.columns:
            df_flat["Prazo"] = df_flat.apply(
                lambda r: _calc_prazo(r.get("card_due"), r.get("card_closed"), report_dt),
                axis=1
            )

        preferred = [
            "card_id", "idShort", "card_name", "list_name", "card_closed", "labels", "members",
            "url", "card_dateLastActivity", "card_start", "card_due", "Prazo",
            "checklist_id", "checklist_name", "checklist_pos",
            "checkitem_id", "checkitem_name", "state", "checkitem_pos", "checkitem_due",
            "responsavel", "responsavel_id",
            "card_desc",
        ]
        cols = [c for c in preferred if c in df_flat.columns] + [c for c in df_flat.columns if c not in preferred]
        df_flat = df_flat[cols]
    else:
        df_flat = pd.DataFrame()

    # ----------------- ChecklistItems (inclui card_due, card_closed e Prazo)
    if not df_checkitems.empty and not df_cards.empty:
        df_checkitems2 = df_checkitems.merge(
            df_cards[["card_id", "card_due", "card_closed", "Prazo"]],
            how="left",
            on="card_id"
        )
        preferred_ci = [
            "card_id", "card_due", "card_closed", "Prazo",
            "checklist_id", "checklist_name",
            "checkitem_id", "checkitem_name", "state", "checkitem_pos", "checkitem_due",
            "responsavel", "responsavel_id",
        ]
        cols = [c for c in preferred_ci if c in df_checkitems2.columns] + [c for c in df_checkitems2.columns if c not in preferred_ci]
        df_checkitems = df_checkitems2[cols]
    else:
        if not df_checkitems.empty and "Prazo" not in df_checkitems.columns:
            df_checkitems["Prazo"] = "Em dia"

    # ----------------- Checklists resumo (opcional)
    if not df_checkitems.empty:
        tmp = df_checkitems.copy()
        tmp["is_complete"] = tmp["state"].fillna("").str.lower().eq("complete")
        grp = tmp.groupby(["card_id", "checklist_id", "checklist_name"], dropna=False).agg(
            total_itens=("checkitem_id", "count"),
            concluidos=("is_complete", "sum"),
        ).reset_index()
        grp["pendentes"] = grp["total_itens"] - grp["concluidos"]
        grp["perc_concluido"] = (grp["concluidos"] / grp["total_itens"]).round(4)
        df_checklists = df_checklists.merge(grp, how="left", on=["card_id", "checklist_id", "checklist_name"])
    else:
        if df_checklists.empty:
            df_checklists = pd.DataFrame(columns=["checklist_id", "card_id", "checklist_name", "checklist_pos"])

    # ----------------- Explore pendências (usa Prazo do card; itens completos fora)
    if not df_flat.empty:
        df_explore = df_flat[df_flat["state"].fillna("").str.lower() != "complete"].copy()
        keep = [
            "list_name", "card_name", "labels", "url", "card_due", "Prazo",
            "checklist_name", "checkitem_name", "responsavel"
        ]
        keep = [c for c in keep if c in df_explore.columns]
        df_explore = df_explore[keep]
        df_explore = df_explore.sort_values(
            by=[c for c in ["list_name", "Prazo", "card_due", "card_name"] if c in df_explore.columns],
            na_position="last"
        )
    else:
        df_explore = pd.DataFrame()

    # Segurança extra (UI/export)
    df_cards = _ensure_unique_columns(df_cards) if not df_cards.empty else df_cards
    df_checklists = _ensure_unique_columns(df_checklists) if not df_checklists.empty else df_checklists
    df_checkitems = _ensure_unique_columns(df_checkitems) if not df_checkitems.empty else df_checkitems
    df_flat = _ensure_unique_columns(df_flat) if not df_flat.empty else df_flat
    df_explore = _ensure_unique_columns(df_explore) if not df_explore.empty else df_explore

    return df_cards, df_checkitems, df_checklists, df_flat, df_explore


# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title="Trello JSON → Excel (Prazo = Concluído)", layout="wide")
st.title("Trello JSON → Excel (Cards + Checklists + Items + FlatExport + Explore)")
st.caption("Coluna 'Prazo': se card concluído => 'Concluído'; caso contrário => 'Em dia'/'Em atraso' pelo due do card.")

uploaded = st.file_uploader("Envie o JSON exportado do Trello", type=["json"])

if uploaded:
    try:
        data = json.load(uploaded)
    except Exception as e:
        st.error(f"Não consegui ler o JSON: {e}")
        st.stop()

    report_dt = datetime.now()

    df_cards, df_checkitems, df_checklists, df_flat, df_explore = parse_trello_export(data, report_dt)

    # -------- Filtro Explore por label
    st.subheader("Explore (pendências) – filtros por label")

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

    # -------- Prévia
    st.subheader("Prévia - Explore (pendências)")
    st.dataframe(df_explore_filtered.head(300), use_container_width=True)

    with st.expander("Prévia - FlatExport"):
        st.dataframe(df_flat.head(200), use_container_width=True)

    with st.expander("Prévia - Cards"):
        st.dataframe(df_cards.head(200), use_container_width=True)

    with st.expander("Prévia - ChecklistItems"):
        st.dataframe(df_checkitems.head(200), use_container_width=True)

    with st.expander("Prévia - Checklists"):
        st.dataframe(df_checklists.head(200), use_container_width=True)

    # -------- Exportação
    st.divider()
    st.subheader("Exportação (.xlsx)")

    try:
        consolidated_bytes = dfs_to_xlsx_bytes(
            {
                "Cards": df_cards,
                "Checklists": df_checklists,
                "ChecklistItems": df_checkitems,
                "FlatExport": df_flat,
                "Explore": df_explore_filtered,
            }
        )

        st.download_button(
            "Baixar Excel consolidado (todas as abas)",
            data=consolidated_bytes,
            file_name="trello_export_consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Baixar Cards.xlsx",
                data=df_to_xlsx_bytes(df_cards, "Cards"),
                file_name="Cards.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.download_button(
                "Baixar Checklists.xlsx",
                data=df_to_xlsx_bytes(df_checklists, "Checklists"),
                file_name="Checklists.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.download_button(
                "Baixar Explore.xlsx (com filtros)",
                data=df_to_xlsx_bytes(df_explore_filtered, "Explore"),
                file_name="Explore.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with c2:
            st.download_button(
                "Baixar ChecklistItems.xlsx",
                data=df_to_xlsx_bytes(df_checkitems, "ChecklistItems"),
                file_name="ChecklistItems.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.download_button(
                "Baixar FlatExport.xlsx",
                data=df_to_xlsx_bytes(df_flat, "FlatExport"),
                file_name="FlatExport.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except ModuleNotFoundError as e:
        st.error(str(e))
        st.info("Inclua `XlsxWriter` (recomendado) ou `openpyxl` no requirements.txt.")
        st.stop()
    except ValueError as e:
        st.error(f"Erro ao exportar Excel: {e}")
        st.stop()

else:
    st.info("Faça upload do JSON do Trello para gerar as visões e exportar o Excel.")
