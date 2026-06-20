# utils/row_generator.py
from .indexer import build_indices

def format_number(value):
    if value is None or value == '':
        return ''
    rounded = round(float(value), 8)
    s = f"{rounded:.8f}".rstrip('0').rstrip('.')
    if '.' not in s:
        s += '.0'
    return s.replace('.', ',')

def generate_vyd_rows(data):
    indices = build_indices(data)
    rows = []
    processed_iv = set()

    for vyd in indices['vyd_list']:
        iv_key = (vyd.get('PLO_NN'), vyd.get('CH_NN'), vyd.get('UCH_NN', 0),
                  vyd.get('CODE'), vyd.get('VN'))
        is_first = iv_key not in processed_iv
        processed_iv.add(iv_key)

        plo = indices['plo'].get(vyd.get('PLO_NN'), {})
        plo_code = plo.get('CODE', '')
        plo_name = plo.get('NAME', '')

        cech = indices['cech'].get((vyd.get('PLO_NN'), vyd.get('CH_NN')), {})
        uch = indices['uch'].get((vyd.get('PLO_NN'), vyd.get('CH_NN'), vyd.get('UCH_NN', 0)), {})

        if is_first:
            cech_code = cech.get('CODE', '')
            cech_name = cech.get('NAME', '')
            uch_code = uch.get('CODE', '')
            uch_name = uch.get('NAME', '')
            try:
                vyd_code = f"{int(vyd.get('CODE', '')):02d}" if vyd.get('CODE') else ''
            except:
                vyd_code = str(vyd.get('CODE', ''))
            vyd_name = vyd.get('NAME', '')
            vyd_vn = vyd.get('VN', '')
            vyd_hourd = vyd.get('HOURD', 0)
            vyd_hours = vyd.get('HOURS', 0)
            vyd_kol = vyd.get('KOL', 1)
        else:
            cech_code = cech_name = uch_code = uch_name = ''
            vyd_code = vyd_name = vyd_vn = ''
            vyd_hourd = vyd_hours = vyd_kol = ''

        vyb_list = indices['vyb_vyd'].get(
            (vyd.get('PLO_NN'), vyd.get('CH_NN'), vyd.get('UCH_NN', 0),
             vyd.get('CODE'), vyd.get('VN')), []
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
        izav_str = ', '.join(izav_parts) if izav_parts else ''

        for idx, (cnt_key, cnt) in enumerate(vyd.get('CNTV', {}).items()):
            is_first_substance = (idx == 0)
            subst_code = cnt.get('CNT_CP')
            subst = indices['subst'].get(subst_code, {})

            gas_set = set()
            for vyb in vyb_list:
                ist_plo = vyb.get('PLO_NNI')
                ist_ch = vyb.get('CH_NNI')
                ist_nn = vyb.get('IST_NN')
                ist_vn = vyb.get('IST_VN')
                gas_val = None
                gasv_key = (vyd.get('PLO_NN'), vyd.get('CH_NN'),
                            vyd.get('CODE'), vyd.get('VN'), subst_code)
                if gasv_key in indices.get('gasv_full', {}):
                    gas_val = indices['gasv_full'][gasv_key].get('GS_KODS', '')
                else:
                    gaso_key = (ist_plo, ist_ch, ist_nn, ist_vn, subst_code)
                    if gaso_key in indices.get('gaso_full', {}):
                        gas_val = indices['gaso_full'][gaso_key].get('GS_KODS', '')
                if gas_val:
                    for part in str(gas_val).split(','):
                        part = part.strip()
                        if part:
                            gas_set.add(part)
            gas_str = ','.join(sorted(gas_set)) if gas_set else ''

            row = {
                'plo_code': plo_code,
                'plo_name': plo_name,
                'col1': cech_code if is_first_substance else '',
                'col2': cech_name if is_first_substance else '',
                'col3': (uch_code if uch_code != 0 else '') if is_first_substance else '',
                'col4': uch_name if is_first_substance else '',
                'col5': vyd_code if is_first_substance else '',
                'col6': vyd_name if is_first_substance else '',
                'col7': vyd_vn if is_first_substance else '',
                'col8': vyd_hourd if is_first_substance else '',
                'col9': vyd_hours if is_first_substance else '',
                'col10': vyd_kol if is_first_substance else '',
                'col11': subst_code,
                'col12': subst.get('NAME', ''),
                'col13': format_number(cnt.get('CNT_GRS', 0)),
                'col14': format_number(cnt.get('CNT_TG', 0)),
                'col15': format_number(cnt.get('CNT_TGTG', 0)),
                'col16': gas_str,
                'col17': izav_str if is_first_substance else '',
                'col18': ''
            }
            rows.append(row)
    return rows