import json
from io import BytesIO
from datetime import datetime
from typing import Any, Dict, Tuple

import pandas as pd
import streamlit as st


# ------------------------- Helpers -------------------------

def _safe_dt(dt_str: str | None):
    """
    Converte ISO do Trello para datetime.
    IMPORTANTE: Excel não suporta datetime com timezone. Então aqui já removemos tzinfo.
    """
    if not dt_str:
        return None
    try:
        # Ex.: "2025-01-01T10:00:00.000Z" -> "+00:00"
        dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        # Remove timezone (tz-aware -> naive), evitando erro no Excel
        if getattr(dt, "tzinfo", None) is not None:
            dt = dt.replace(tzinfo=None)
        return dt
    except Exception:
        return dt_str  # fallback: mantém string


def _label_display(lbl: dict) -> str:
    name = (lbl.get("name") or "").strip()
    if name:
        return name
    color = (lbl.get("color") or "").strip()
    return f"(label:{color})" if color else "(label)"


def _ensure_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Evita colunas duplicadas (Streamlit/PyArrow quebra)."""
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


def _pick_excel_engine() -> str:
    """
    Seleciona um engine disponível:
    - xlsxwriter (preferido)
    - openpyxl (fallback)
    """
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
    """
    Sanitiza valores para Excel:
    - datetime tz-aware -> naive
    - pandas Timestamp tz-aware -> naive
    - dict/list -> JSON string
    """
    # pandas Timestamp
    if isinstance(v, pd.Timestamp):
        if v.tz is not None:
            return v.tz_convert(None).to_pydatetime()
        return v.to_pydatetime()

    # python datetime
    if isinstance(v, datetime):
        if v.tzinfo is not None:
            return v.replace(tzinfo=None)
        return v

    # dict/list (evita exportação problemática)
    if isinstance(v, (dict, list)):
        try:
            return json.dumps(v, ensure_ascii=False)
        except Exception:
            return str(v)

    return v


def sanitize_df_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica sanitização em todas as células, garantindo compatibilidade com Excel."""
    if df.empty:
        return df
    out = df.copy()

    # Aplica sanitização célula a célula
    out = out.map(_sanitize_value_for_excel)

    # Garante colunas únicas também (segurança extra)
    out = _ensure_unique_columns(out)
    return out


def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    engine = _pick_excel_engine()
    output = BytesIO()
    df2 = sanitize_df_for_excel(df)

    with pd.ExcelWriter(output, engine=engine) as writer:
        df2.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return output.getvalue()


def dfs_to_xlsx_bytes(dfs: Dict[str, pd.DataFrame]) -> bytes:
    engine = _pick_excel_engine()
    output = BytesIO()

    with pd.ExcelWriter(output, engine=engine) as writer:
        for name, df in dfs.items():
            df2 = sanitize_df_for_excel(df)
            df2.to_excel(writer, index=False, sheet_name=name[:31])

    return output.getvalue()


# ------------------------- Parser -------------------------

def parse_trello_export(data: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    lists_map = {l.get("id"): l.get("name") for l in (data.get("lists", []) or [])}
    members_map = {
        m.get("id"): (m.get("fullName") or m.get("username") or m.get("id"))
        for m in (data.get("members", []) or [])
    }
    labels_map = {lb.get("id"): _label_display(lb) for lb in (data.get("labels", []) or [])}

    # Checklists e itens
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

    # Cards
    card_rows = []
    for c in (data.get("cards", []) or []):
        card_id = c.get("id")
        id_list = c.get("idList")

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

    # Agregado de checklist por card (texto)
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

    # FlatExport
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

    # Segurança extra contra colunas duplicadas antes da UI
    df_cards = _ensure_unique_columns(df_cards)
    df_checklists = _ensure_unique_columns(df_checklists)
    df_checkitems = _ensure_unique_columns(df_checkitems)
    df_flat = _ensure_unique_columns(df_flat)

    return df_cards, df_checkitems, df_checklists, df_flat


# ------------------------- Streamlit UI -------------------------

st.set_page_config(page_title="Trello JSON → Excel", layout="wide")
st.title("Trello JSON → Excel (Cards + Checklists + FlatExport)")

st.caption("Exporta um Excel consolidado (com abas) e também arquivos XLSX individuais por tabela.")

uploaded = st.file_uploader("Envie o JSON exportado do Trello", type=["json"])

if uploaded:
    try:
        data = json.load(uploaded)
    except Exception as e:
        st.error(f"Não consegui ler o JSON: {e}")
        st.stop()

    df_cards, df_checkitems, df_checklists, df_flat = parse_trello_export(data)

    # Prévia (não sanitiza aqui; já garantimos colunas únicas)
    st.subheader("Prévia - FlatExport (1 linha por item de checklist)")
    st.dataframe(df_flat.head(200), use_container_width=True)

    with st.expander("Prévia - Cards"):
        st.dataframe(df_cards.head(200), use_container_width=True)

    with st.expander("Prévia - ChecklistItems"):
        st.dataframe(df_checkitems.head(200), use_container_width=True)

    with st.expander("Prévia - Checklists"):
        st.dataframe(df_checklists.head(200), use_container_width=True)

    st.divider()
    st.subheader("Exportação")

    # Consolidado
    try:
        all_bytes = dfs_to_xlsx_bytes(
            {
                "Cards": df_cards,
                "Checklists": df_checklists,
                "ChecklistItems": df_checkitems,
                "FlatExport": df_flat,
            }
        )
        st.download_button(
            "Baixar Excel consolidado (todas as abas)",
            data=all_bytes,
            file_name="trello_export_consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except ModuleNotFoundError as e:
        st.error(str(e))
        st.info("Solução: adicione `XlsxWriter` (recomendado) ou `openpyxl` no requirements.txt.")
        st.stop()
    except ValueError as e:
        st.error(f"Erro ao exportar Excel: {e}")
        st.info("Este app já sanitiza timezone em datetimes; se persistir, existe algum tipo de dado não suportado em alguma coluna.")
        st.stop()

    # Individuais
    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "Baixar Cards.xlsx",
            data=df_to_xlsx_bytes(df_cards, sheet_name="Cards"),
            file_name="Cards.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "Baixar Checklists.xlsx",
            data=df_to_xlsx_bytes(df_checklists, sheet_name="Checklists"),
            file_name="Checklists.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with col2:
        st.download_button(
            "Baixar ChecklistItems.xlsx",
            data=df_to_xlsx_bytes(df_checkitems, sheet_name="ChecklistItems"),
            file_name="ChecklistItems.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "Baixar FlatExport.xlsx",
            data=df_to_xlsx_bytes(df_flat, sheet_name="FlatExport"),
            file_name="FlatExport.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    st.info("Faça upload do JSON do Trello para gerar as tabelas e baixar os XLSX.")
