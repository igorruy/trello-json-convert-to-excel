# -------- FlatExport (1 linha por item de checklist, com dados do card)
if not df_checkitems.empty and not df_cards.empty:
    # Renomeia colunas que conflitam ANTES do merge
    df_checkitems_flat = df_checkitems.rename(
        columns={
            "due": "checkitem_due",
            "pos": "checkitem_pos",
        }
    )

    df_cards_flat = df_cards.rename(
        columns={
            "due": "card_due",
            "start": "card_start",
            "dateLastActivity": "card_dateLastActivity",
            "closed": "card_closed",
            "desc": "card_desc",
        }
    )

    df_flat = df_checkitems_flat.merge(
        df_cards_flat,
        how="left",
        on="card_id",
    )

    # Reordena colunas (opcional)
    preferred = [
        "card_id", "idShort", "card_name", "list_name", "card_closed", "labels", "members",
        "url", "card_dateLastActivity", "card_start", "card_due",
        "checklist_id", "checklist_name",
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
            "checklist_id", "checklist_name",
            "checkitem_id", "checkitem_name", "state", "checkitem_pos", "checkitem_due",
            "card_desc",
        ]
    )
