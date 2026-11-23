import os
import re
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, Reference
from openpyxl.formatting.rule import ColorScaleRule

def _score_sheet(ws):
    rows = list(ws.iter_rows(values_only=True))
    cnt = 0
    for r in rows:
        if sum(1 for c in r if c is not None and str(c).strip() != '') >= 3:
            cnt += 1
    return cnt

def _detect_header_row(rows):
    best_idx = 0
    best_score = -1
    limit = min(30, len(rows))
    for i in range(limit):
        r = rows[i]
        score = sum(1 for c in r if c is not None and str(c).strip() != '')
        # prefer rows with string labels
        str_score = sum(1 for c in r if isinstance(c, str) and c.strip() != '')
        total = score + str_score
        if total > best_score:
            best_score = total
            best_idx = i
    return best_idx

def load_workbook_df(path):
    wb = openpyxl.load_workbook(path)
    if 'Data_Transaksi' in wb.sheetnames:
        sheet_name = 'Data_Transaksi'
    else:
        sheet_name = max(wb.sheetnames, key=lambda s: _score_sheet(wb[s]))
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Empty sheet")
    # cached values (data_only) for formula cells
    try:
        wbv = openpyxl.load_workbook(path, data_only=True)
        ws_v = wbv[sheet_name]
        rows_v = list(ws_v.iter_rows(values_only=True))
    except Exception:
        rows_v = rows
    if len(rows_v) != len(rows):
        rows_v = rows
    hi = _detect_header_row(rows)
    header = [str(c).strip() if c is not None else f"Col{i}" for i, c in enumerate(rows[hi])]
    # dedup empty header names
    seen = {}
    for i, h in enumerate(header):
        if h == '' or h.startswith('Col') and h[3:].isdigit() and rows[hi][i] is None:
            header[i] = f"Col{i}"
        if header[i] in seen:
            seen[header[i]] += 1
            header[i] = f"{header[i]}_{seen[header[i]]}"
        else:
            seen[header[i]] = 0
    raw_data = rows[hi+1:]
    val_data = rows_v[hi+1:]
    # filter empty rows based on formula rows
    filtered = [(r, rv) for r, rv in zip(raw_data, val_data) if any(c is not None and str(c).strip() != '' for c in r)]
    raw_data = [x[0] for x in filtered]
    val_data = [x[1] for x in filtered]
    ncols = len(header)
    data = []
    for r, rv in zip(raw_data, val_data):
        rr = list(r) + [None]*(ncols-len(r)) if len(r) < ncols else list(r[:ncols])
        vv = list(rv) + [None]*(ncols-len(rv)) if len(rv) < ncols else list(rv[:ncols])
        new_row = []
        for cf, cv in zip(rr, vv):
            if isinstance(cf, str) and cf.strip().startswith('='):
                if isinstance(cv, str) and cv.strip().startswith('#'):
                    new_row.append(None)
                elif cv is None:
                    new_row.append(cf)
                else:
                    new_row.append(cv)
            else:
                new_row.append(cf)
        data.append(new_row)
    df = pd.DataFrame(data, columns=header)
    return wb, ws, df, header, sheet_name

def is_formula(v):
    return isinstance(v, str) and v.strip().startswith('=')

def non_formula_cols(df):
    cols = []
    for c in df.columns:
        if not any(is_formula(v) for v in df[c].tolist()):
            cols.append(c)
    if not cols:
        cols = list(df.columns)
    return cols

def empty_mask(df, cols=None):
    if cols is None:
        cols = list(df.columns)
    mask = pd.DataFrame(False, index=df.index, columns=cols)
    for c in cols:
        s = df[c]
        def is_empty(x):
            if x is None:
                return True
            if isinstance(x, float) and pd.isna(x):
                return True
            if isinstance(x, str) and x.strip() == '':
                return True
            return False
        mask[c] = s.map(is_empty)
    return mask.any(axis=1)

def detect_duplicates(df):
    nf = non_formula_cols(df)
    dup_mask = df[nf].duplicated(keep='first')
    return df[dup_mask].copy(), nf

def detect_empty(df, cols=None):
    m = empty_mask(df, cols)
    return df[m].copy()

def clean_data(df, sort_by=None, group_by=None, empty_cols=None):
    nf = non_formula_cols(df)
    if empty_cols is None:
        empty_cols = [c for c in df.columns if not empty_mask(df, [c]).all()]
        if not empty_cols:
            empty_cols = list(df.columns)
    m = empty_mask(df, empty_cols)
    df2 = df[~m].copy()
    df2 = df2.drop_duplicates(subset=nf, keep='first')
    if sort_by and sort_by in df2.columns:
        df2 = df2.sort_values(by=sort_by, kind='stable')
    elif group_by and group_by in df2.columns:
        df2 = df2.sort_values(by=group_by, kind='stable')
    return df2

def sanitize_sheet_name(name):
    name = re.sub(r'[\[\]:*?/\\]', '_', str(name))
    name = name.strip()
    if not name:
        name = 'Sheet'
    return name[:31]

def write_df_to_sheet(ws, df, header=None, start_row=1, start_col=1, style_header=True):
    if header is None:
        header = list(df.columns)
    for j, h in enumerate(header, start=start_col):
        cell = ws.cell(row=start_row, column=j, value=h)
        if style_header:
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='4472C4')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    r = start_row + 1
    for _, row in df.iterrows():
        for j, h in enumerate(header, start=start_col):
            v = row[h]
            if isinstance(v, float) and pd.isna(v):
                v = None
            ws.cell(row=r, column=j, value=v)
        r += 1
    for j, h in enumerate(header, start=start_col):
        ws.column_dimensions[get_column_letter(j)].width = min(28, max(10, len(str(h)) + 2))
    ws.freeze_panes = ws.cell(row=start_row+1, column=start_col)
    return ws

def add_category_sheets(wb, df, col):
    created = []
    if col not in df.columns:
        return created
    for val in list(df[col].dropna().unique()):
        sub = df[df[col] == val].copy()
        name = sanitize_sheet_name(f"Cat_{val}")
        base = name
        i = 1
        while name in wb.sheetnames:
            name = sanitize_sheet_name(f"{base}_{i}")
            i += 1
        ws = wb.create_sheet(title=name)
        write_df_to_sheet(ws, sub)
        created.append(name)
    return created

def make_chart_sheet(wb, df, selected, options=None):
    options = options or {}
    name = 'Charts'
    i = 1
    while name in wb.sheetnames:
        name = f"Charts_{i}"
        i += 1
    ws = wb.create_sheet(title=name)
    row_cursor = 1
    cat_col = options.get('cat_col')
    if not cat_col:
        for c in df.columns:
            if (str(df[c].dtype) in ('object','string','str') or df[c].dtype == object) and df[c].nunique() <= 15 and not any(is_formula(v) for v in df[c]):
                cat_col = c
                break
    if cat_col is None:
        cat_col = df.columns[0]
    num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if not num_cols:
        for c in df.columns:
            try:
                pd.to_numeric(df[c])
                num_cols.append(c)
            except Exception:
                pass
    if 'Bar' in selected:
        counts = df[cat_col].value_counts().sort_index()
        ws.cell(row=row_cursor, column=1, value='Bar Chart Data').font = Font(bold=True)
        row_cursor += 1
        ws.cell(row=row_cursor, column=1, value=cat_col)
        ws.cell(row=row_cursor, column=2, value='Count')
        hr = row_cursor
        row_cursor += 1
        for k, v in counts.items():
            ws.cell(row=row_cursor, column=1, value=str(k))
            ws.cell(row=row_cursor, column=2, value=int(v))
            row_cursor += 1
        end = row_cursor - 1
        chart = BarChart()
        chart.title = f"Bar Chart - Count by {cat_col}"
        chart.y_axis.title = 'Count'
        chart.x_axis.title = cat_col
        data = Reference(ws, min_col=2, min_row=hr, max_row=end)
        cats = Reference(ws, min_col=1, min_row=hr+1, max_row=end)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, f"D{hr}")
        row_cursor = end + 15
    if 'Pie' in selected:
        counts = df[cat_col].value_counts().sort_index()
        ws.cell(row=row_cursor, column=1, value='Pie Chart Data').font = Font(bold=True)
        row_cursor += 1
        ws.cell(row=row_cursor, column=1, value=cat_col)
        ws.cell(row=row_cursor, column=2, value='Count')
        hr = row_cursor
        row_cursor += 1
        for k, v in counts.items():
            ws.cell(row=row_cursor, column=1, value=str(k))
            ws.cell(row=row_cursor, column=2, value=int(v))
            row_cursor += 1
        end = row_cursor - 1
        chart = PieChart()
        chart.title = f"Pie Chart - {cat_col}"
        data = Reference(ws, min_col=2, min_row=hr, max_row=end)
        cats = Reference(ws, min_col=1, min_row=hr+1, max_row=end)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, f"D{hr}")
        row_cursor = end + 15
    if 'Line' in selected and num_cols:
        metric = num_cols[0]
        agg = df.groupby(cat_col)[metric].mean().sort_index()
        ws.cell(row=row_cursor, column=1, value='Line Chart Data').font = Font(bold=True)
        row_cursor += 1
        ws.cell(row=row_cursor, column=1, value=cat_col)
        ws.cell(row=row_cursor, column=2, value=f'Mean {metric}')
        hr = row_cursor
        row_cursor += 1
        for k, v in agg.items():
            ws.cell(row=row_cursor, column=1, value=str(k))
            ws.cell(row=row_cursor, column=2, value=float(v) if pd.notna(v) else None)
            row_cursor += 1
        end = row_cursor - 1
        chart = LineChart()
        chart.title = f"Line Chart - Mean {metric} by {cat_col}"
        chart.y_axis.title = metric
        chart.x_axis.title = cat_col
        data = Reference(ws, min_col=2, min_row=hr, max_row=end)
        cats = Reference(ws, min_col=1, min_row=hr+1, max_row=end)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, f"D{hr}")
        row_cursor = end + 15
    if 'Histogram' in selected and num_cols:
        metric = num_cols[0]
        s = pd.to_numeric(df[metric], errors='coerce').dropna()
        if len(s) == 0:
            return name
        try:
            arr_counts, arr_bins = np.histogram(s, bins=10)
        except Exception:
            arr_counts, arr_bins = np.histogram(s, bins='auto')
        ws.cell(row=row_cursor, column=1, value='Histogram Data').font = Font(bold=True)
        row_cursor += 1
        ws.cell(row=row_cursor, column=1, value='Bin')
        ws.cell(row=row_cursor, column=2, value='Count')
        hr = row_cursor
        row_cursor += 1
        for i in range(len(arr_counts)):
            lo = arr_bins[i]
            hi = arr_bins[i+1]
            ws.cell(row=row_cursor, column=1, value=f"{lo:.2f} - {hi:.2f}")
            ws.cell(row=row_cursor, column=2, value=int(arr_counts[i]))
            row_cursor += 1
        end = row_cursor - 1
        chart = BarChart()
        chart.type = 'col'
        chart.title = f"Histogram - {metric}"
        chart.y_axis.title = 'Frequency'
        chart.x_axis.title = metric
        chart.gapWidth = 0
        data = Reference(ws, min_col=2, min_row=hr, max_row=end)
        cats = Reference(ws, min_col=1, min_row=hr+1, max_row=end)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, f"D{hr}")
        row_cursor = end + 15
    if 'Stacked' in selected:
        cat_cols = [c for c in df.columns if (str(df[c].dtype) in ('object','string','str') or df[c].dtype == object) and df[c].nunique() <= 15 and c != cat_col and not any(is_formula(v) for v in df[c])]
        second = cat_cols[0] if cat_cols else None
        if second:
            ct = pd.crosstab(df[cat_col], df[second])
            ws.cell(row=row_cursor, column=1, value='Stacked Chart Data').font = Font(bold=True)
            row_cursor += 1
            ws.cell(row=row_cursor, column=1, value=cat_col)
            for j, colname in enumerate(ct.columns, start=2):
                ws.cell(row=row_cursor, column=j, value=str(colname))
            hr = row_cursor
            row_cursor += 1
            for idxname, rrow in ct.iterrows():
                ws.cell(row=row_cursor, column=1, value=str(idxname))
                for j, colname in enumerate(ct.columns, start=2):
                    ws.cell(row=row_cursor, column=j, value=int(rrow[colname]))
                row_cursor += 1
            end = row_cursor - 1
            chart = BarChart()
            chart.type = 'col'
            chart.grouping = 'stacked'
            chart.title = f"Stacked Chart - {cat_col} by {second}"
            data = Reference(ws, min_col=2, max_col=1+len(ct.columns), min_row=hr, max_row=end)
            cats = Reference(ws, min_col=1, min_row=hr+1, max_row=end)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            ws.add_chart(chart, f"D{hr}")
            row_cursor = end + 15
    if 'Scatter' in selected and len(num_cols) >= 2:
        xcol, ycol = num_cols[0], num_cols[1]
        sub = df[[xcol, ycol]].dropna().head(200)
        ws.cell(row=row_cursor, column=1, value='Scatter Data').font = Font(bold=True)
        row_cursor += 1
        ws.cell(row=row_cursor, column=1, value=xcol)
        ws.cell(row=row_cursor, column=2, value=ycol)
        hr = row_cursor
        row_cursor += 1
        for _, rrow in sub.iterrows():
            ws.cell(row=row_cursor, column=1, value=float(rrow[xcol]))
            ws.cell(row=row_cursor, column=2, value=float(rrow[ycol]))
            row_cursor += 1
        end = row_cursor - 1
        chart = ScatterChart()
        chart.title = f"Scatter - {xcol} vs {ycol}"
        chart.x_axis.title = xcol
        chart.y_axis.title = ycol
        xvalues = Reference(ws, min_col=1, min_row=hr+1, max_row=end)
        yvalues = Reference(ws, min_col=2, min_row=hr+1, max_row=end)
        from openpyxl.chart import Series
        series = Series(yvalues, xvalues, title=f"{ycol}")
        chart.series.append(series)
        ws.add_chart(chart, f"D{hr}")
        row_cursor = end + 15
    if 'Heatmap' in selected:
        cat_cols = [c for c in df.columns if (str(df[c].dtype) in ('object','string','str') or df[c].dtype == object) and df[c].nunique() <= 15 and not any(is_formula(v) for v in df[c])]
        if len(cat_cols) >= 2:
            c1, c2 = cat_cols[0], cat_cols[1]
        elif len(cat_cols) == 1:
            c1, c2 = cat_cols[0], cat_cols[0]
        else:
            c1, c2 = df.columns[0], df.columns[1]
        ct = pd.crosstab(df[c1], df[c2])
        ws.cell(row=row_cursor, column=1, value=f'Heatmap - {c1} x {c2}').font = Font(bold=True)
        row_cursor += 1
        ws.cell(row=row_cursor, column=1, value=c1 + '\\' + c2)
        for j, colname in enumerate(ct.columns, start=2):
            ws.cell(row=row_cursor, column=j, value=str(colname))
        hr = row_cursor
        row_cursor += 1
        for idxname, rrow in ct.iterrows():
            ws.cell(row=row_cursor, column=1, value=str(idxname))
            for j, colname in enumerate(ct.columns, start=2):
                ws.cell(row=row_cursor, column=j, value=int(rrow[colname]))
            row_cursor += 1
        end = row_cursor - 1
        data_range = f"B{hr+1}:{get_column_letter(1+len(ct.columns))}{end}"
        ws.conditional_formatting.add(data_range, ColorScaleRule(start_type='min', start_color='FFFFFF',
            mid_type='percentile', mid_value=50, mid_color='FFC7CE',
            end_type='max', end_color='C00000'))
        row_cursor = end + 2
    if 'Funnel' in selected:
        counts = df[cat_col].value_counts().sort_values(ascending=False)
        ws.cell(row=row_cursor, column=1, value='Funnel Chart Data').font = Font(bold=True)
        row_cursor += 1
        ws.cell(row=row_cursor, column=1, value=cat_col)
        ws.cell(row=row_cursor, column=2, value='Count')
        hr = row_cursor
        row_cursor += 1
        for k, v in counts.items():
            ws.cell(row=row_cursor, column=1, value=str(k))
            ws.cell(row=row_cursor, column=2, value=int(v))
            row_cursor += 1
        end = row_cursor - 1
        chart = BarChart()
        chart.type = 'bar'
        chart.title = f"Funnel Chart - {cat_col}"
        data = Reference(ws, min_col=2, min_row=hr, max_row=end)
        cats = Reference(ws, min_col=1, min_row=hr+1, max_row=end)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, f"D{hr}")
        row_cursor = end + 15
    return name

def process_file(input_path, output_path, options, progress_cb=None):
    def prog(msg, pct):
        if progress_cb:
            progress_cb(msg, pct)
    prog("Loading workbook...", 5)
    wb, ws, df, header, sheet_name = load_workbook_df(input_path)
    prog("Detecting duplicates...", 15)
    dup_df, nf = detect_duplicates(df)
    prog("Detecting empty rows...", 25)
    if options.get('empty_cols'):
        empty_cols = options.get('empty_cols')
    else:
        # skip columns that are entirely empty (e.g. converted error formula columns)
        empty_cols = []
        for c in df.columns:
            if not empty_mask(df, [c]).all():
                empty_cols.append(c)
        if not empty_cols:
            empty_cols = list(df.columns)
    empty_df = detect_empty(df, empty_cols)
    prog("Creating duplicate sheet...", 35)
    if options.get('add_dup_sheet', True) and len(dup_df) > 0:
        name = 'Data_Duplikat'
        i = 1
        while name in wb.sheetnames:
            name = f'Data_Duplikat_{i}'; i += 1
        wsd = wb.create_sheet(title=name)
        write_df_to_sheet(wsd, dup_df)
    prog("Creating empty sheet...", 45)
    if options.get('add_empty_sheet', True) and len(empty_df) > 0:
        name = 'Data_Kosong'
        i = 1
        while name in wb.sheetnames:
            name = f'Data_Kosong_{i}'; i += 1
        wse = wb.create_sheet(title=name)
        write_df_to_sheet(wse, empty_df)
    prog("Cleaning data...", 55)
    clean_df = None
    if options.get('add_clean_sheet', True):
        sort_by = options.get('sort_by')
        group_by = options.get('group_by')
        clean_df = clean_data(df, sort_by=sort_by, group_by=group_by, empty_cols=empty_cols)
        name = 'Data_Bersih'
        i = 1
        while name in wb.sheetnames:
            name = f'Data_Bersih_{i}'; i += 1
        wsc = wb.create_sheet(title=name)
        write_df_to_sheet(wsc, clean_df)
        if options.get('add_category_sheets', True):
            col = options.get('category_col')
            if not col:
                for c in df.columns:
                    if (str(df[c].dtype) in ('object','string','str') or df[c].dtype == object) and df[c].nunique() <= 15 and not any(is_formula(v) for v in df[c]):
                        col = c; break
            if col:
                src = options.get('category_source', 'clean')
                if src == 'clean' and clean_df is not None and len(clean_df) > 0:
                    cat_base = clean_df
                elif src == 'filtered':
                    ec = empty_cols if empty_cols is not None else ([c for c in df.columns if not empty_mask(df, [c]).all()] or list(df.columns))
                    m = empty_mask(df, ec)
                    tmp = df[~m].copy()
                    nf = non_formula_cols(df)
                    tmp = tmp.drop_duplicates(subset=nf, keep='first')
                    cat_base = tmp
                else:
                    cat_base = df
                prog("Creating category sheets...", 65)
                add_category_sheets(wb, cat_base, col)
        if options.get('add_charts', True):
            prog("Creating charts...", 75)
            selected = options.get('chart_types', [])
            if selected:
                make_chart_sheet(wb, clean_df if len(clean_df)>0 else df, selected, {'cat_col': options.get('chart_cat_col') or col})
    prog("Saving...", 90)
    wb.save(output_path)
    prog("Done", 100)
    summary = {
        'rows': len(df),
        'duplicates': len(dup_df),
        'empty': len(empty_df),
        'clean': len(clean_df) if clean_df is not None else 0,
        'output': output_path,
    }
    return summary
