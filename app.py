import json
from io import BytesIO
from datetime import datetime
from typing import Any, Dict, Tuple

import pandas as pd
import streamlit as st


# ------------------------- Excel / Sanitização -------------------------

def _pick_excel_engine() -> str:
    """Seleciona um engine disponível para gerar XLSX."""
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
    """Evita erro do Excel com tz-aware datetime e valores complexos."""
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

    # dict/list -> string
    if isinstance(v, (dict, list)):
        try:
            return json.dumps(v, ensure_ascii=False)
        except Exception:
            return str(v)

    return v


def sanitize_df_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
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
            df2.to_excel(writer, index=False, sheet_name=sheet[:31])
    return output.getvalue()


def df_to_xlsx_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    return dfs_to_xlsx_bytes({sheet_name: df})


# ------------------------- Helpers Trello -------------------------

def _safe_dt(dt_str: str | None):
    """
    Converte ISO do Trello para datetime.
    Remove tzinfo para não quebrar exportação no Excel.
    """
    if not dt_str:
        return None
    try:
        dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        if dt.tzinfo is not None:
            dt = dt.replace(tzinfo=None)
        return dt
    except Exception:
        return dt_str


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


# ------------------------- Parser -------------------------

def parse_trello_export(data: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Retorna:
      - df_geral: visão completa (1 linha por item de checklist + dados do card)
      - df_explore: pendências (state != complete) com responsável (idMember) e labels do card
    """
    # Mapas auxiliares
    lists_map = {l.get("id"): l.get("name") for l in (data.get("lists", []) or [])}
    members_map = {
        m.get("id"): (m.get("fullName") or m.get("username") or m.get("id"))
        for m in (data.get("members", []) or [])
    }
    labels_map = {lb.get("id"): _label_display(lb) for lb in (data.get("labels", []) or [])}

    # Cards -> tabela base
    card_rows = []
    for c in (data.get("cards", []) or []):
        card_id = c.get("id")
        id_list = c.get("idList")

        # Labels do card
        card_labels = []
        for lb in (c.get("labels", []) or []):
            lb_id = lb.get("id")
            if lb_id and lb_id in labels_map:
                card_labels.append(labels_map[lb_id])
            else:
                card_labels.append(_label_display(lb))
        card_labels = sorted({x for x in card_labels if x})

        # Membros do card (pode ser útil como contexto)
        card_members = [members_map.get(mid, mid) for mid in (c.get("idMembers", []) or [])]

        card_rows.append(
            {
                "card_id": card_id,
                "idShort": c.get("idShort"),
                "card_name": c.get("name"),
                "list_name": lists_map.get(id_list),
                "card_closed": c.get("closed"),
                "labels": ", ".join(card_labels),
                "members": ", ".join(card_members),
                "url": c.get("url"),
                "shortLink": c.get("shortLink"),
                "card_due": _safe_dt(c.get("due")),
                "card_start": _safe_dt(c.get("start")),
                "card_dateLastActivity": _safe_dt(c.get("dateLastActivity")),
                "card_desc": c.get("desc"),
            }
        )

    df_cards = pd.DataFrame(card_rows)

    # Checklists/itens -> 1 linha por checkitem
    checkitem_rows = []
    for cl in (data.get("checklists", []) or []):
        cl_id = cl.get("id")
        card_id = cl.get("idCard")
        cl_name = cl.get("name")
        cl_pos = cl.get("pos")

        for it in (cl.get("checkItems", []) or []):
            id_member = it.get("idMember")
            resp = members_map.get(id_member) if id_member else None

            checkitem_rows.append(
                {
                    "card_id": card_id,
                    "checklist_id": cl_id,
                    "checklist_name": cl_name,
                    "checklist_pos": cl_pos,
                    "checkitem_id": it.get("id"),
                    "checkitem_name": it.get("name"),
                    "state": it.get("state"),
                    "checkitem_pos": it.get("pos"),
                    "checkitem_due": _safe_dt(it.get("due")),
                    "responsavel_id": id_member,
                    "responsavel": resp,
                }
            )

    df_items = pd.DataFrame(checkitem_rows)

    # ---- Aba "Geral": conteúdo completo atual (FlatExport-like)
    if not df_items.empty and not df_cards.empty:
        df_geral = df_items.merge(df_cards, how="left", on="card_id")
    elif not df_cards.empty:
        # Caso sem checklists: ainda exporta cards, mas mantém colunas de itens vazias
        base_cols = [
            "card_id", "checklist_id", "checklist_name", "checklist_pos",
            "checkitem_id", "checkitem_name", "state", "checkitem_pos", "checkitem_due",
            "responsavel_id", "responsavel",
        ]
        df_geral = pd.DataFrame(columns=base_cols).merge(df_cards, how="right", on="card_id")
    else:
        df_geral = pd.DataFrame()

    # Ordenação "boa" para Geral
    if not df_geral.empty:
        preferred = [
            "card_id", "idShort", "card_name", "list_name", "card_closed", "labels", "members",
            "url", "card_dateLastActivity", "card_start", "card_due",
            "checklist_id", "checklist_name", "checklist_pos",
            "checkitem_id", "checkitem_name", "state", "checkitem_pos", "checkitem_due",
            "responsavel", "responsavel_id",
            "card_desc",
        ]
        cols = [c for c in preferred if c in df_geral.columns] + [c for c in df_geral.columns if c not in preferred]
        df_geral = df_geral[cols]

    # ---- Aba "Explore": pendências (state != complete)
    if not df_geral.empty:
        df_explore = df_geral.copy()
        df_explore = df_explore[df_explore["state"].fillna("").str.lower() != "complete"].copy()

        # Campos gerenciais: deixar mais “enxuto”
        keep = [
            "card_id", "idShort", "card_name", "list_name", "labels", "url",
            "checklist_name", "checkitem_name", "state",
            "responsavel", "checkitem_due",
            "card_dateLastActivity", "card_due",
        ]
        keep = [c for c in keep if c in df_explore.columns]
        df_explore = df_explore[keep]

        # Ordenação: por card, checklist, prazo
        sort_cols = [c for c in ["list_name", "card_name", "checklist_name", "checkitem_due"] if c in df_explore.columns]
        if sort_cols:
            df_explore = df_explore.sort_values(by=sort_cols, na_position="last")
    else:
        df_explore = pd.DataFrame()

    # Garantir colunas únicas para UI/export
    df_geral = _ensure_unique_columns(df_geral) if not df_geral.empty else df_geral
    df_explore = _ensure_unique_columns(df_explore) if not df_explore.empty else df_explore

    return df_geral, df_explore


# ------------------------- Streamlit UI -------------------------

st.set_page_config(page_title="Trello JSON → Excel (Geral + Explore)", layout="wide")
st.title("Trello JSON → Excel")
st.caption("Gera 2 abas: 'Geral' (conteúdo completo) e 'Explore' (pendências gerenciais com filtros por label).")

uploaded = st.file_uploader("Envie o JSON exportado do Trello", type=["json"])

if uploaded:
    try:
        data = json.load(uploaded)
    except Exception as e:
        st.error(f"Não consegui ler o JSON: {e}")
        st.stop()

    df_geral, df_explore = parse_trello_export(data)

    # ---------------- Filters (Explore) ----------------
    st.subheader("Explore: Pendências (itens de checklist não concluídos)")

    all_labels = []
    if not df_explore.empty and "labels" in df_explore.columns:
        # labels é string "A, B, C"
        uniq = set()
        for s in df_explore["labels"].fillna("").astype(str).tolist():
            for part in [p.strip() for p in s.split(",") if p.strip()]:
                uniq.add(part)
        all_labels = sorted(uniq)

    selected_labels = st.multiselect(
        "Filtrar por label (mostra cards que contenham pelo menos uma das labels selecionadas)",
        options=all_labels,
        default=[],
    )

    df_explore_filtered = df_explore.copy()
    if selected_labels and not df_explore_filtered.empty and "labels" in df_explore_filtered.columns:
        def has_any_label(label_str: str) -> bool:
            ls = {p.strip() for p in str(label_str).split(",") if p.strip()}
            return any(l in ls for l in selected_labels)

        df_explore_filtered = df_explore_filtered[df_explore_filtered["labels"].apply(has_any_label)].copy()

    # Prévia
    st.dataframe(df_explore_filtered.head(300), use_container_width=True)

    with st.expander("Prévia - Geral (conteúdo completo)"):
        st.dataframe(df_geral.head(300), use_container_width=True)

    # ---------------- Exportação ----------------
    st.divider()
    st.subheader("Exportação (.xlsx)")

    try:
        # Consolidado (2 abas)
        consolidated_bytes = dfs_to_xlsx_bytes(
            {
                "Geral": df_geral,
                "Explore": df_explore_filtered,  # exporta já com filtro aplicado
            }
        )

        st.download_button(
            "Baixar Excel consolidado (Geral + Explore)",
            data=consolidated_bytes,
            file_name="trello_export_geral_explore.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "Baixar Geral.xlsx",
                data=df_to_xlsx_bytes(df_geral, "Geral"),
                file_name="Geral.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with col2:
            st.download_button(
                "Baixar Explore.xlsx (com filtros)",
                data=df_to_xlsx_bytes(df_explore_filtered, "Explore"),
                file_name="Explore.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except ModuleNotFoundError as e:
        st.error(str(e))
        st.info("Inclua `XlsxWriter` (recomendado) ou `openpyxl` no requirements.txt.")
        st.stop()
    except ValueError as e:
        st.error(f"Erro ao exportar Excel: {e}")
        st.info("Se persistir, pode haver algum valor em formato não suportado em alguma coluna. Eu ajusto a sanitização se você colar o detalhe do log.")
        st.stop()

else:
    st.info("Faça upload do JSON do Trello para gerar as abas 'Geral' e 'Explore'.")
