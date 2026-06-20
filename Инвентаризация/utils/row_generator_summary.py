# utils/row_generator_summary.py
from .indexer import build_indices

def format_number(value, decimals=8):
    """Форматирует число: 8 знаков после запятой, запятая, 0,0 для нуля."""
    if value is None or value == '':
        return '0,0'
    try:
        num = float(value)
    except:
        return str(value)
    rounded = round(num, decimals)
    s = f"{rounded:.{decimals}f}".replace('.', ',')
    if abs(rounded) < 1e-9:
        return '0,0'
    return s

def generate_summary_rows(data):
    indices = build_indices(data)
    subst_data = {}

    # --- 1. Проход по VYD (источники выделения) ---
    for vyd in indices['vyd_list']:
        plo_nn = vyd.get('PLO_NN')
        ch_nn = vyd.get('CH_NN')
        uch_nn = vyd.get('UCH_NN', 0)
        vyd_code = vyd.get('CODE')
        vyd_vn = vyd.get('VN')
        for cnt in vyd.get('CNTV', {}).values():
            cnt_cp = cnt.get('CNT_CP')
            if cnt_cp is None:
                continue
            tgtg = cnt.get('CNT_TGTG', 0.0)   # валовый выброс от ИВ (до очистки)
            if cnt_cp not in subst_data:
                subst_data[cnt_cp] = {
                    'total_from_vyd': 0.0,
                    'total_without_cleanup': 0.0,
                    'total_without_cleanup_org': 0.0,
                    'total_to_cleanup': 0.0,
                    'total_after_cleanup': 0.0,
                }
            subst_data[cnt_cp]['total_from_vyd'] += tgtg

            # Проверяем, есть ли газоочистка для этого ИВ и вещества
            has_gasv = (plo_nn, ch_nn, vyd_code, vyd_vn, cnt_cp) in indices.get('gasv_full', {})
            has_gaso = False
            vyb_list = indices['vyb_vyd'].get((plo_nn, ch_nn, uch_nn, vyd_code, vyd_vn), [])
            for vyb in vyb_list:
                ist_plo = vyb.get('PLO_NNI')
                ist_ch = vyb.get('CH_NNI')
                ist_nn = vyb.get('IST_NN')
                ist_vn = vyb.get('IST_VN')
                if (ist_plo, ist_ch, ist_nn, ist_vn, cnt_cp) in indices.get('gaso_full', {}):
                    has_gaso = True
                    break

            if has_gasv or has_gaso:
                # Выброс направляется на очистку
                subst_data[cnt_cp]['total_to_cleanup'] += tgtg
                # Если есть GASV, сразу берём очищенное значение (CNT_TG)
                if has_gasv:
                    gasv = indices['gasv_full'].get((plo_nn, ch_nn, vyd_code, vyd_vn, cnt_cp), {})
                    cleaned = gasv.get('CNT_TG', 0.0)
                    subst_data[cnt_cp]['total_after_cleanup'] += cleaned
            else:
                # Нет очистки – распределяем выброс по долям (K_VYD_IST)
                if not vyb_list:
                    subst_data[cnt_cp]['total_without_cleanup'] += tgtg
                else:
                    total_share = 0.0
                    for vyb in vyb_list:
                        share = vyb.get('K_VYD_IST', 0) / 100.0
                        total_share += share
                        ist_plo = vyb.get('PLO_NNI')
                        ist_ch = vyb.get('CH_NNI')
                        ist_nn = vyb.get('IST_NN')
                        ist_vn = vyb.get('IST_VN')
                        ist = indices['ist'].get((ist_plo, ist_ch, ist_nn, ist_vn))
                        if ist and ist.get('IST_TP') in (1, 4, 6, 7, 9, 10):
                            subst_data[cnt_cp]['total_without_cleanup_org'] += tgtg * share
                        else:
                            subst_data[cnt_cp]['total_without_cleanup'] += tgtg * share
                    if total_share < 0.999:
                        subst_data[cnt_cp]['total_without_cleanup'] += tgtg * (1 - total_share)

    # --- 2. Проход по IST (источники выброса) для организованных газоочисток ---
    # Добавляем очищенные выбросы для веществ, которые прошли через GASO
    for ist in indices.get('ist_list', []):
        plo_nn = ist.get('PLO_NN')
        ch_nn = ist.get('CH_NN')
        ist_nn = ist.get('IST_NN')
        ist_vn = ist.get('IST_VN')
        for cnt in ist.get('CNT', {}).values():
            cnt_cp = cnt.get('CNT_CP')
            if cnt_cp is None:
                continue
            key = (plo_nn, ch_nn, ist_nn, ist_vn, cnt_cp)
            if key in indices.get('gaso_full', {}):
                # Выброс после очистки берём из IST.CNT.CNT_TG
                cleaned = cnt.get('CNT_TG', 0.0)
                if cnt_cp not in subst_data:
                    subst_data[cnt_cp] = {
                        'total_from_vyd': 0.0,
                        'total_without_cleanup': 0.0,
                        'total_without_cleanup_org': 0.0,
                        'total_to_cleanup': 0.0,
                        'total_after_cleanup': 0.0,
                    }
                subst_data[cnt_cp]['total_after_cleanup'] += cleaned

    # --- Формирование строк с разделением на твёрдые и жидкие/газообразные ---
    subst_info = {s['CODE']: s.get('AGR', 1) for s in indices['subst_list']}
    rows_solid = []
    rows_liquid = []

    for cnt_cp, data in subst_data.items():
        ag = subst_info.get(cnt_cp, 1)
        total_from_vyd = data['total_from_vyd']
        without = data['total_without_cleanup']
        without_org = data['total_without_cleanup_org']
        to_cleanup = data['total_to_cleanup']
        after_cleanup = data['total_after_cleanup']
        captured = to_cleanup - after_cleanup
        utilized = 0.0
        total_emitted = without + after_cleanup

        row = {
            'type': 'data',
            'code': cnt_cp,
            'name': indices['subst'].get(cnt_cp, {}).get('NAME', ''),
            'col3': total_from_vyd,
            'col4': without,
            'col5': without_org,
            'col6': to_cleanup,
            'col7': captured,
            'col8': utilized,
            'col9': after_cleanup,
            'col10': total_emitted,
        }
        if ag == 1:
            rows_solid.append(row)
        else:
            rows_liquid.append(row)

    rows_solid.sort(key=lambda x: x['code'])
    rows_liquid.sort(key=lambda x: x['code'])

    final_rows = []
    if rows_solid:
        final_rows.append({'type': 'header', 'title': 'Загрязняющие вещества - твердые :'})
        final_rows.extend(rows_solid)
    if rows_liquid:
        final_rows.append({'type': 'header', 'title': 'Загрязняющие вещества - жидкие и газообразные :'})
        final_rows.extend(rows_liquid)

    return final_rows, subst_info