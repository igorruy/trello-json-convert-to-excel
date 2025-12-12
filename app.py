import json
from io import BytesIO
from datetime import datetime, date
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
    raise ModuleNotFoundError("Nenhum engine de Excel disponível. Instale XlsxWriter ou openpyxl.")


def _sanitize_value(v: Any) -> Any:
    if isinstance(v, pd.Timestamp):
        return v.tz_convert(None).to_pydatetime() if v.tz else v.to_pydatetime()
    if isinstance(v, datetime):
        return v.replace(tzinfo=None) if v.tzinfo else v
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
            # garante que a aba exista mesmo vazia
            if df2 is None or df2.empty:
                pd.DataFrame({"info": ["Sem dados"]}).to_excel(writer, sheet_name=sheet[:31], index=False)
            else:
                df2.to_excel(writer, sheet_name=sheet[:31], index=False)

    return output.getvalue()


# =========================
# Helpers – Trello
# =========================

def _safe_dt(val):
    if not val:
        return None
    try:
        # Trello costuma trazer Z (UTC). Removemos tzinfo para Excel.
        return datetime.fromisoformat(str(val).replace("Z", "+00:00")).replace(tzinfo=None)
    except Exception:
        return None


def _label_display(lbl):
    name = (lbl.get("name") or "").strip()
    if name:
        return name
    color = (lbl.get("color") or "").strip()
    return f"(label:{color})" if color else "(label)"


def _calc_prazo(card_due: datetime | None, state: str | None, completed_at: datetime | None, report_dt: datetime) -> str:
    """
    Prazo:
      - deadline = card_due
      - se concluído: compara completed_at (quando existir) vs card_due; senão usa report_dt
      - se pendente: compara report_dt vs card_due
    """
    if not card_due:
        return "Em dia"  # sem prazo definido no card -> não marca atraso

    state_norm = (state or "").strip().lower()
    ref_dt = report_dt

    if state_norm == "complete":
        ref_dt = completed_at or report_dt

    # compara por data (não por hora)
    return "Em atraso" if ref_dt.date() > card_due.date() else "Em dia"


# =========================
# Parser principal
# =========================

def parse_trello(data: dict, report_dt: datetime) -> Tuple[pd.DataFrame, pd.DataFrame]:
    lists = {l.get("id"): l.get("name") for l in data.get("lists", []) or []}
    members = {
        m.get("id"): (m.get("fullName") or m.get("username") or m.get("id"))
        for m in data.get("members", []) or []
    }
    labels_map = {l.get("id"): _label_display(l) for l in data.get("labels", []) or []}

    # ---- Cards
    cards = []
    for c in data.get("cards", []) or []:
        card_id = c.get("id")
        id_list = c.get("idList")

        card_labels = []
        for lb in c.get("labels", []) or []:
            lb_id = lb.get("id")
            card_labels.append(labels_map.get(lb_id, _label_display(lb)))

        card_due = _safe_dt(c.get("due"))

        cards.append({
            "card_id": card_id,
            "card_name": c.get("name"),
            "list_name": lists.get(id_list),
            "labels": ", ".join(sorted(set([x for x in card_labels if x]))),
            "members": ", ".join(members.get(mid, mid) for mid in (c.get("idMembers", []) or [])),
            "card_due": card_due,
            "url": c.get("url"),
        })

    df_cards = pd.DataFrame(cards)

    # ---- Checklist items
    items = []
    for cl in data.get("checklists", []) or []:
        card_id = cl.get("idCard")
        cl_name = cl.get("name")

        for it in cl.get("checkItems", []) or []:
            state = it.get("state")

            # Tentativa de capturar data de conclusão (varia conforme export)
            completed_at = _safe_dt(
                it.get("dateCompleted")
                or it.get("dateComplete")
                or it.get("completedAt")
                or it.get("dateCompletion")
            )

            # responsável do item (quando existir)
            id_member = it.get("idMember")
            resp = members.get(id_member) if id_member else None

            items.append({
                "card_id": card_id,
                "checklist_name": cl_name,
                "checkitem_name": it.get("name"),
                "state": state,
                "responsavel": resp,
                "checkitem_completed_at": completed_at,
            })

    df_items = pd.DataFrame(items)

    # ---- ABA GERAL (SEM FILTRO)
    if not df_items.empty and not df_cards.empty:
        df_geral = df_items.merge(df_cards, on="card_id", how="left")
    elif not df_cards.empty:
        df_geral = df_cards.copy()
    else:
        df_geral = pd.DataFrame()

    # ---- Adiciona Prazo na Geral (para linhas de checklist e também cards puros)
    if not df_geral.empty:
        if "card_due" not in df_geral.columns:
            df_geral["card_due"] = None

        # Para linhas sem item (caso não tenha checklist): state/completed_at podem não existir
        if "state" not in df_geral.columns:
            df_geral["state"] = None
        if "checkitem_completed_at" not in df_geral.columns:
            df_geral["checkitem_completed_at"] = None

        df_geral["Prazo"] = df_geral.apply(
            lambda r: _calc_prazo(
                r.get("card_due"),
                r.get("state"),
                r.get("checkitem_completed_at"),
                report_dt
            ),
            axis=1
        )

    # ---- ABA EXPLORE (pendências)
    if not df_geral.empty and "state" in df_geral.columns:
        df_explore = df_geral[df_geral["state"].fillna("").str.lower() != "complete"].copy()
    else:
        df_explore = pd.DataFrame(columns=df_geral.columns if not df_geral.empty else [])

    # Mantém visão gerencial enxuta
    if not df_explore.empty:
        cols = [
            "list_name", "card_name", "labels", "url", "card_due",
            "checklist_name", "checkitem_name", "responsavel",
            "Prazo"
        ]
        cols = [c for c in cols if c in df_explore.columns]
        df_explore = df_explore[cols]

        df_explore = df_explore.sort_values(
            by=[c for c in ["list_name", "card_name", "Prazo", "card_due"] if c in df_explore.columns],
            na_position="last"
        )

    return df_geral, df_explore


# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title="Trello → Excel (Geral / Explore)", layout="wide")
st.title("Trello JSON → Excel")
st.caption("Abas: Geral (tudo) | Explore (pendências gerenciais). Inclui coluna 'Prazo' (Em dia / Em atraso).")

uploaded = st.file_uploader("Envie o JSON exportado do Trello", type=["json"])

if uploaded:
    data = json.load(uploaded)

    # Data/hora de geração do relatório (local do servidor; suficiente para regra Em dia/Em atraso)
    report_dt = datetime.now()

    df_geral, df_explore = parse_trello(data, report_dt)

    # -------- Filtro Explore por label
    st.subheader("Explore – Pendências (itens de checklist não concluídos)")

    labels = []
    if not df_explore.empty and "labels" in df_explore.columns:
        uniq = set()
        for s in df_explore["labels"].fillna("").astype(str).tolist():
            for part in [p.strip() for p in s.split(",") if p.strip()]:
                uniq.add(part)
        labels = sorted(uniq)

    selected = st.multiselect("Filtrar por label", labels, default=[])

    df_explore_f = df_explore.copy()
    if selected and not df_explore_f.empty and "labels" in df_explore_f.columns:
        df_explore_f = df_explore_f[
            df_explore_f["labels"].fillna("").astype(str).apply(
                lambda x: any(l in {p.strip() for p in x.split(",") if p.strip()} for l in selected)
            )
        ].copy()

    st.dataframe(df_explore_f, use_container_width=True)

    with st.expander("Visualizar aba Geral (conteúdo completo)"):
        st.dataframe(df_geral, use_container_width=True)

    # -------- Exportação
    st.divider()
    st.subheader("Exportação (.xlsx)")

    try:
        excel_bytes = export_excel({
            "Geral": df_geral,         # SEM filtro
            "Explore": df_explore_f,   # COM filtro aplicado
        })

        st.download_button(
            "Baixar Excel (Geral + Explore)",
            data=excel_bytes,
            file_name="trello_geral_explore.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except ModuleNotFoundError as e:
        st.error(str(e))
        st.info("Inclua `XlsxWriter` (recomendado) ou `openpyxl` no requirements.txt.")
        st.stop()

else:
    st.info("Faça upload do JSON do Trello para gerar as abas 'Geral' e 'Explore'.")
