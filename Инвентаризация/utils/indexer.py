# utils/indexer.py
def build_indices(data):
    """Строит индексы для быстрого доступа к данным JSON."""
    indices = {}

    # Индекс площадок по коду
    indices['plo'] = {}
    for key, plo in data.get('PLO', {}).items():
        code = plo.get('CODE')
        if code is not None:
            indices['plo'][code] = plo

    # Индекс цехов по (площадка, код_цеха)
    indices['cech'] = {}
    for key, cech in data.get('CECH', {}).items():
        plo = cech.get('PLO_NN')
        ch_code = cech.get('CODE')
        indices['cech'][(plo, ch_code)] = cech

    # Индекс участков по (площадка, цех, код_участка)
    indices['uch'] = {}
    for key, uch in data.get('UCH', {}).items():
        plo = uch.get('PLO_NN')
        ch = uch.get('CH_NN')
        code = uch.get('CODE')
        indices['uch'][(plo, ch, code)] = uch

    # Индекс веществ по коду
    indices['subst'] = {}
    for key, subst in data.get('ShortSubst', {}).items():
        code = subst.get('CODE')
        indices['subst'][code] = subst

    # Список всех веществ (для перебора)
    indices['subst_list'] = list(data.get('ShortSubst', {}).values())

    # Список всех источников выделения (VYD)
    indices['vyd_list'] = list(data.get('VYD', {}).values())

    # Список всех ИЗАВ (IST)
    indices['ist_list'] = list(data.get('IST', {}).values())

    # --- Индексы для связей VYB_VYD ---
    indices['vyb_vyd'] = {}
    for key, vyb in data.get('VYB_VYD', {}).items():
        plo = vyb.get('PLO_NNV')
        ch = vyb.get('CH_NNV')
        uch = vyb.get('UCH_NN', 0)
        vyd_nn = vyb.get('VYD_NN')
        vyd_vn = vyb.get('VYD_VN')
        key_tuple = (plo, ch, uch, vyd_nn, vyd_vn)
        indices['vyb_vyd'].setdefault(key_tuple, []).append(vyb)

    # --- Индекс IST ---
    indices['ist'] = {}
    for key, ist in data.get('IST', {}).items():
        plo = ist.get('PLO_NN')
        ch = ist.get('CH_NN')
        ist_nn = ist.get('IST_NN')
        ist_vn = ist.get('IST_VN')
        indices['ist'][(plo, ch, ist_nn, ist_vn)] = ist

    # --- GASV (полные записи) ---
    indices['gasv_full'] = {}
    for key, gasv in data.get('GASV', {}).items():
        plo = gasv.get('PLO_NN')
        ch = gasv.get('CH_NN')
        vyd_nn = gasv.get('VYD_NN')
        vyd_vn = gasv.get('VYD_VN')
        cnt_cp = gasv.get('CNT_CP')
        if gasv.get('VIRT') == 1 and not gasv.get('GS_KODS'):
            continue
        key_tuple = (plo, ch, vyd_nn, vyd_vn, cnt_cp)
        indices['gasv_full'][key_tuple] = gasv

        # GASO (организованные газоочистки) – полные записи (включая виртуальные)
    indices['gaso_full'] = {}
    for key, gaso in data.get('GASO', {}).items():
        plo = gaso.get('PLO_NN')
        ch = gaso.get('CH_NN')
        ist_nn = gaso.get('IST_NN')
        ist_vn = gaso.get('IST_VN')
        cnt_cp = gaso.get('CNT_CP')
        key_tuple = (plo, ch, ist_nn, ist_vn, cnt_cp)
        indices['gaso_full'][key_tuple] = gaso

    # Метеопараметры
    indices['meteo'] = data.get('METEO', {})

    return indices
