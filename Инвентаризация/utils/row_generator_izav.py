# utils/row_generator_izav.py
from .indexer import build_indices

def format_number(value, decimals=8, zero_as_empty=False):
    if value is None or value == '':
        return ''
    try:
        num = float(value)
    except:
        return str(value)
    if zero_as_empty and abs(num) < 1e-9:
        return ''
    rounded = round(num, decimals)
    s = f"{rounded:.{decimals}f}".rstrip('0').rstrip('.')
    if '.' not in s:
        s += '.0'
    return s.replace('.', ',')

def format_int(value, zero_as_empty=False):
    if value is None or value == '':
        return ''
    try:
        num = int(round(float(value)))
    except:
        return str(value)
    if zero_as_empty and num == 0:
        return ''
    return str(num)

def format_3dec(value, zero_as_empty=False):
    return format_number(value, 3, zero_as_empty)

def format_5dec(value, zero_as_empty=False):
    return format_number(value, 5, zero_as_empty)

def generate_izav_rows(data):
    indices = build_indices(data)
    rows = []

    ist_list = indices.get('ist_list', [])
    ist_list.sort(key=lambda x: (x.get('PLO_NN', 0), x.get('CH_NN', 0), x.get('XCODE', ''), x.get('IST_VN', 0)))

    total_by_key = {}
    for ist in ist_list:
        xcode = ist.get('XCODE')
        if not xcode:
            continue
        for cnt in ist.get('CNT', {}).values():
            cnt_cp = cnt.get('CNT_CP')
            if cnt_cp is None:
                continue
            tgtg = cnt.get('CNT_TGTG', 0)
            key = (xcode, cnt_cp)
            total_by_key[key] = total_by_key.get(key, 0) + tgtg

    first_variant_for_subst = {}
    for ist in ist_list:
        xcode = ist.get('XCODE')
        if not xcode:
            continue
        ist_vn = ist.get('IST_VN', 1)
        for cnt in ist.get('CNT', {}).values():
            cnt_cp = cnt.get('CNT_CP')
            if cnt_cp is None:
                continue
            key = (xcode, cnt_cp)
            if key not in first_variant_for_subst or ist_vn < first_variant_for_subst[key]:
                first_variant_for_subst[key] = ist_vn

    current_plo_cech = None
    for ist in ist_list:
        plo_nn = ist.get('PLO_NN')
        ch_nn = ist.get('CH_NN')
        plo = indices['plo'].get(plo_nn, {})
        cech = indices['cech'].get((plo_nn, ch_nn), {}) if ch_nn else {}
        plo_code = plo.get('CODE', '')
        plo_name = plo.get('NAME', '')
        cech_code = cech.get('CODE', '')
        cech_name = cech.get('NAME', '')
        plo_cech_key = (plo_nn, ch_nn)
        if plo_cech_key != current_plo_cech:
            rows.append({'type': 'header', 'plo_code': plo_code, 'plo_name': plo_name,
                         'cech_code': cech_code, 'cech_name': cech_name})
            current_plo_cech = plo_cech_key

        xcode = ist.get('XCODE', '')
        ist_vn = ist.get('IST_VN', 1)

        ist_tp = ist.get('IST_TP', 0)
        if ist_tp == 1:
            type_str = "Организованный"
        elif ist_tp == 3:
            type_str = "Неорганизованный"
        elif ist_tp == 4:
            type_str = "Совокупность точечных"
        else:
            type_str = ""

        height = format_int(ist.get('IST_H', 0), zero_as_empty=False)
        diam = format_3dec(ist.get('IST_D', 0), zero_as_empty=True)
        width = ""
        length = ""
        x1 = format_int(ist.get('IST_X1', 0), zero_as_empty=False)
        y1 = format_int(ist.get('IST_Y1', 0), zero_as_empty=False)
        x2 = format_int(ist.get('IST_X2', 0), zero_as_empty=False) if ist.get('IST_X2') else ''
        y2 = format_int(ist.get('IST_Y2', 0), zero_as_empty=False) if ist.get('IST_Y2') else ''
        z = format_int(ist.get('IST_Z', 0), zero_as_empty=True)
        mode = ist_vn
        speed = format_3dec(ist.get('IST_W', 0), zero_as_empty=True)
        volume = format_3dec(ist.get('IST_V', 0), zero_as_empty=True)
        temp = format_int(ist.get('IST_T', 0), zero_as_empty=True)
        density = format_3dec(ist.get('IST_P', indices['meteo'].get('P', 1.29)), zero_as_empty=True)

        cnt_items = list(ist.get('CNT', {}).items())
        for idx, (_, cnt) in enumerate(cnt_items):
            cnt_cp = cnt.get('CNT_CP')
            if cnt_cp is None:
                continue
            subst = indices['subst'].get(cnt_cp, {})
            is_first_in_variant = (idx == 0)
            key = (xcode, cnt_cp)
            first_vn = first_variant_for_subst.get(key, 999)
            show_total = (ist_vn == first_vn)
            total_t = format_3dec(total_by_key.get(key, 0), zero_as_empty=False) if show_total else ''

            row = {
                'type': 'data',
                'col1': xcode if is_first_in_variant else '',
                'col2': type_str if is_first_in_variant else '',
                'col3': ist.get('IST_NAME', '') if is_first_in_variant else '',
                'col4': ist.get('IST_KOL', 1) if is_first_in_variant else '',
                'col5': height if is_first_in_variant else '',
                'col6': diam if is_first_in_variant else '',
                'col7': width,
                'col8': length,
                'col9': x1 if is_first_in_variant else '',
                'col10': y1 if is_first_in_variant else '',
                'col11': x2 if is_first_in_variant else '',
                'col12': y2 if is_first_in_variant else '',
                'col13': z if is_first_in_variant else '',
                'col14': mode if is_first_in_variant else '',
                'col15': speed if is_first_in_variant else '',
                'col16': speed if is_first_in_variant else '',
                'col17': volume if is_first_in_variant else '',
                'col18': temp if is_first_in_variant else '',
                'col19': density if is_first_in_variant else '',
                'col20': cnt_cp,
                'col21': subst.get('NAME', ''),
                'col22': format_3dec(cnt.get('CNT_GM3', 0), zero_as_empty=False),
                'col23': format_3dec(cnt.get('CNT_GRS', 0), zero_as_empty=False),
                'col24': format_3dec(cnt.get('CNT_TG', 0), zero_as_empty=False),
                'col25': total_t,
                'col26': ''
            }
            rows.append(row)
    return rows