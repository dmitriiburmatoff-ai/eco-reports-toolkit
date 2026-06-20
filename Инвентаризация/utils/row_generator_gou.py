# utils/row_generator_gou.py
from .indexer import build_indices

def format_value(value, decimals=2):
    if value is None or value == '':
        return ''
    try:
        num = float(value)
    except:
        return str(value)
    rounded = round(num, decimals)
    s = f"{rounded:.{decimals}f}".rstrip('0').rstrip('.')
    if '.' not in s:
        s += '.0'
    return s.replace('.', ',')

def format_iv_code(code):
    try:
        return f"{int(code):02d}"
    except:
        return str(code)

def generate_gou_rows(data):
    indices = build_indices(data)
    rows = []

    # --- 1. Строки по ИВ (через GASV) ---
    for vyd in indices['vyd_list']:
        plo_nn = vyd.get('PLO_NN')
        ch_nn = vyd.get('CH_NN')
        uch_nn = vyd.get('UCH_NN', 0)
        plo = indices['plo'].get(plo_nn, {})
        cech = indices['cech'].get((plo_nn, ch_nn), {})
        uch = indices['uch'].get((plo_nn, ch_nn, uch_nn), {})
        plo_code = plo.get('CODE', '')
        plo_name = plo.get('NAME', '')
        cech_code = cech.get('CODE', '')
        cech_name = cech.get('NAME', '')
        uch_code = uch.get('CODE', '') if uch_nn != 0 else ''
        uch_name = uch.get('NAME', '') if uch_nn != 0 else ''
        vyd_code_fmt = format_iv_code(vyd.get('CODE', ''))
        source_name = f"ИВ: {vyd.get('NAME', '')} ({vyd_code_fmt})"

        # Находим связанные ИЗАВ через VYB_VYD
        vyb_list = indices['vyb_vyd'].get(
            (plo_nn, ch_nn, uch_nn, vyd.get('CODE'), vyd.get('VN')), []
        )
        izav_parts = []
        for vyb in vyb_list:
            ist_plo = vyb.get('PLO_NNI')
            ist_ch = vyb.get('CH_NNI')
            ist_nn = vyb.get('IST_NN')
            ist_vn = vyb.get('IST_VN')
            ist = indices['ist'].get((ist_plo, ist_ch, ist_nn, ist_vn))
            if ist and ist.get('XCODE'):
                xcode = ist['XCODE']
                if ist_vn != 1:
                    izav_parts.append(f"{xcode}({ist_vn})")
                else:
                    izav_parts.append(xcode)
        izav_codes_str = ', '.join(sorted(izav_parts))

        for cnt in vyd.get('CNTV', {}).values():
            cnt_cp = cnt.get('CNT_CP')
            subst = indices['subst'].get(cnt_cp, {})
            gasv_key = (plo_nn, ch_nn, vyd.get('CODE'), vyd.get('VN'), cnt_cp)
            if gasv_key in indices['gasv_full']:
                gasv = indices['gasv_full'][gasv_key]
                rows.append({
                    'plo_code': plo_code,
                    'plo_name': plo_name,
                    'cech_code': cech_code,
                    'cech_name': cech_name,
                    'uch_code': uch_code,
                    'uch_name': uch_name,
                    'source_name': source_name,
                    'gas_name': f"{gasv.get('GS_NAME', '')} ({gasv.get('GS_KODS', '')})",
                    'izav': izav_codes_str,
                    'eff_pr': format_value(gasv.get('GS_MAX', 0), 2),
                    'eff_fact': format_value(gasv.get('GS_SR', 0), 2),
                    'subst_name': f"{subst.get('NAME', '')} ({cnt_cp})",
                    'coef_norm': format_value(gasv.get('GS_OBESN', 0), 2),
                    'coef_fact': format_value(gasv.get('GS_OBESF', 0), 2),
                })

    # --- 2. Строки по ИЗАВ (через GASO) ---
    for ist in indices.get('ist_list', []):
        plo_nn = ist.get('PLO_NN')
        ch_nn = ist.get('CH_NN')
        plo = indices['plo'].get(plo_nn, {})
        cech = indices['cech'].get((plo_nn, ch_nn), {})
        plo_code = plo.get('CODE', '')
        plo_name = plo.get('NAME', '')
        cech_code = cech.get('CODE', '')
        cech_name = cech.get('NAME', '')
        xcode = ist.get('XCODE', '')
        ist_vn = ist.get('IST_VN', 1)
        xcode_display = f"{xcode}({ist_vn})" if ist_vn != 1 else xcode
        source_name = f"ИЗАВ: {ist.get('IST_NAME', '')} ({xcode})"

        for cnt in ist.get('CNT', {}).values():
            cnt_cp = cnt.get('CNT_CP')
            subst = indices['subst'].get(cnt_cp, {})
            gaso_key = (plo_nn, ch_nn, ist.get('IST_NN'), ist.get('IST_VN'), cnt_cp)
            if gaso_key in indices['gaso_full']:
                gaso = indices['gaso_full'][gaso_key]
                rows.append({
                    'plo_code': plo_code,
                    'plo_name': plo_name,
                    'cech_code': cech_code,
                    'cech_name': cech_name,
                    'uch_code': '',
                    'uch_name': '',
                    'source_name': source_name,
                    'gas_name': f"{gaso.get('GS_NAME', '')} ({gaso.get('GS_KODS', '')})",
                    'izav': xcode_display,
                    'eff_pr': format_value(gaso.get('GS_MAX', 0), 2),
                    'eff_fact': format_value(gaso.get('GS_SR', 0), 2),
                    'subst_name': f"{subst.get('NAME', '')} ({cnt_cp})",
                    'coef_norm': format_value(gaso.get('GS_OBESN', 0), 2),
                    'coef_fact': format_value(gaso.get('GS_OBESF', 0), 2),
                })

    # Сортировка по площадке, цеху, затем по типу (сначала ИВ, потом ИЗАВ) – можно просто по цеху
    rows.sort(key=lambda x: (x['plo_code'], x['cech_code'], x['source_name']))
    return rows