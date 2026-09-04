import io, sys, contextlib, json, datetime, os


def run_generate(input_path, params_json):
    import build_construction_log as bcl
    params = json.loads(params_json)
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            (meta, items, CATEGORIES, N_ITEMS, grand_total, records, grand_total_source, self_sum,
             used_sheet) = bcl.parse_pcces_file(
                input_path, sheet_name=(params.get('sheet') or None),
                agency_override=(params.get('agency') or None),
                project_name_override=(params.get('project_name') or None),
                location_override=(params.get('location') or None),
                proj_code_override=(params.get('proj_code') or None))

            start_date = (datetime.datetime.strptime(params['start_date'], '%Y-%m-%d').date()
                          if params.get('start_date') else datetime.date.today())
            days, _end_date, days_source = bcl.resolve_days(
                start_date,
                days_arg=(int(params['days']) if params.get('days') else None),
                end_date_arg=(params.get('end_date') or None))
            n_mat = int(params.get('n_mat') or 12)
            n_labor = int(params.get('n_labor') or 8)
            n_mach = int(params.get('n_mach') or 8)

            wb = bcl.openpyxl.Workbook()
            ctx = bcl.build_day_database(wb, meta, items, CATEGORIES, N_ITEMS, grand_total,
                                          start_date, days, n_mat, n_labor, n_mach)
            ctx = bcl.build_sched_and_gantt(wb, meta, ctx)
            bcl.build_daily_report(wb, meta, ctx)
            bcl.build_dashboard_sheet(wb, meta, ctx)
            bcl.build_source_ref_sheet(wb, meta, input_path, used_sheet, records, items, CATEGORIES,
                                        N_ITEMS, grand_total, grand_total_source, self_sum)
            safe_name = bcl.re.sub(r'[\\/:*?"<>|]', '_', meta['project_name'] or 'CONSTRUCTION_LOG')
            out_name = (params.get('out_name') or '').strip() or f'{safe_name}_施工日誌.xlsx'
            if not out_name.lower().endswith('.xlsx'):
                out_name += '.xlsx'
            out_path = '/tmp/' + out_name
            wb.save(out_path)
            meta_out = dict(ctx)
            meta_out['CATEGORIES'] = [[label, sn, en, tp, cl]
                                       for label, sn, en, tp, cl, c1, c2 in ctx['CATEGORIES_full']]
            del meta_out['CATEGORIES_full']
            meta_out.update(project_name=meta['project_name'], location=meta['location'],
                             proj_code=meta['proj_code'], agency=meta['agency'], grand_total=grand_total)
            meta_out['start_date'] = str(start_date)
            meta_json_str = json.dumps(meta_out, ensure_ascii=False, indent=2, default=str)

            print(f'工期天數依據：{days_source}')
            print(f'完成！已輸出：{out_name}')
        with open(out_path, 'rb') as f:
            out_bytes = f.read()

        return {
            'ok': True,
            'log': buf.getvalue(),
            'out_name': out_name,
            'out_bytes': out_bytes,
            'meta_json': meta_json_str,
            'meta_name': os.path.splitext(out_name)[0] + '_meta.json',
        }
    except Exception as e:
        import traceback
        return {
            'ok': False,
            'log': buf.getvalue(),
            'error': str(e),
            'traceback': traceback.format_exc(),
        }


def run_merge_change_order(old_xlsx_path, new_pcces_path, params_json):
    """把「變更後」的PCCES檔案(new_pcces_path)套用到既有、已經填了逐日資料的施工日誌
    (old_xlsx_path)，輸出一份全新檔案，原始舊檔案不會被修改。對應GUI的「套用契約變更」模式。"""
    import build_construction_log as bcl
    params = json.loads(params_json)
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            wb, meta, ctx, extra = bcl.apply_change_order(
                old_xlsx_path, new_pcces_path,
                sheet_name=(params.get('sheet') or None),
                agency_override=(params.get('agency') or None),
                project_name_override=(params.get('project_name') or None),
                location_override=(params.get('location') or None),
                proj_code_override=(params.get('proj_code') or None),
                end_date=(params.get('end_date') or None),
                days_override=(int(params['days']) if params.get('days') else None),
                n_mat=(int(params['n_mat']) if params.get('n_mat') else None),
                n_labor=(int(params['n_labor']) if params.get('n_labor') else None),
                n_mach=(int(params['n_mach']) if params.get('n_mach') else None),
            )
            safe_name = bcl.re.sub(r'[\\/:*?"<>|]', '_', meta['project_name'] or 'CONSTRUCTION_LOG')
            out_name = (params.get('out_name') or '').strip() or f'{safe_name}_施工日誌(契約變更).xlsx'
            if not out_name.lower().endswith('.xlsx'):
                out_name += '.xlsx'
            out_path = '/tmp/' + out_name
            wb.save(out_path)
            print(f'完成！契約變更已套用，已輸出全新檔案：{out_name}（原始舊檔案未被修改）')
            for w in extra['warnings']:
                print(f'ℹ️ {w}')
        with open(out_path, 'rb') as f:
            out_bytes = f.read()

        return {
            'ok': True,
            'log': buf.getvalue(),
            'out_name': out_name,
            'out_bytes': out_bytes,
            'unmatched_old_items': extra['unmatched_old_items'],
            'unmatched_new_items': extra['unmatched_new_items'],
            'changed_items': extra['changed_items'],
            'warnings': extra['warnings'],
            'old_days': extra['old_days'],
            'new_days': extra['new_days'],
        }
    except Exception as e:
        import traceback
        return {
            'ok': False,
            'log': buf.getvalue(),
            'error': str(e),
            'traceback': traceback.format_exc(),
        }
