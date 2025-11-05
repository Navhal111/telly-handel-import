"""
Combined Excel processor for files that contain both Attendance and Payroll tables in a single sheet.

This module provides a helper class CombinedProcessor which attempts to detect two header rows
(one for attendance, one for payroll) inside the same sheet, slices the sheet into two tables,
parses them into structured Python dicts and returns both results.

It is deliberately conservative and forgiving: it uses keyword heuristics to detect headers and
parses numeric columns by stripping commas and whitespace. The output is intended to be passed
on to `TallyApiService` or shown to the user.

Place this file in `src/` and import it from your GUI or scripts like:

from src.combined_processor import CombinedProcessor

result = CombinedProcessor().process_combined_sheet(path_to_excel)

"""
from typing import Dict, Any, List, Optional, Tuple
import os
import pandas as pd
import re


def _normalize_header(h: str) -> str:
    """Normalize a header string to a compact lowercase key."""
    if h is None:
        return ""
    s = str(h).strip()
    # remove repeated spaces and non-alphanum
    s = re.sub(r"[\s\u00A0]+", " ", s)
    s = re.sub(r"[^0-9A-Za-z ]", "", s)
    return s.strip().lower()


def _parse_number(v: object) -> Optional[float]:
    """Try to convert common number strings like '1,54,000' or ' 22,00,000 ' to float.
    Returns None if conversion fails or value is blank.
    """
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s == "":
        return None
    # remove commas and spaces
    s = s.replace(',', '').replace(' ', '')
    # sometimes hyphens are used for empty
    if s in ['-', '—', '–']:
        return None
    try:
        return float(s)
    except Exception:
        return None


class CombinedProcessor:
    """Processor to extract both attendance and payroll tables from a single Excel sheet.

    Usage:
        cp = CombinedProcessor()
        result = cp.process_combined_sheet('file.xlsx')

    The returned dict contains keys:
        - attendance_result: dict or None
        - payroll_result: dict or None
    """

    def __init__(self):
        # Heuristics keywords
        self.att_keywords = ['attendance', 'attendance/production', 'attendance days']
        self.pay_keywords = ['basic', 'gross', 'net', 'total', 'salary']
        self.empl_keys = ['empl', 'empl no', 'empl.no', 's.no']

    def _find_header_indices(self, df: pd.DataFrame) -> Tuple[Optional[int], Optional[int]]:
        """Scan rows to find indices of attendance and payroll header rows.

        Returns (attendance_idx, payroll_idx)
        """
        attendance_idx = None
        payroll_idx = None

        for idx, row in df.iterrows():
            # build uppercase joined string of row
            row_values = [str(x) for x in row.values if not pd.isna(x)]
            if not row_values:
                continue
            row_text = ' '.join(row_values).upper()

            # detect attendance header
            if attendance_idx is None:
                if any(k.upper() in row_text for k in self.att_keywords) and any(k.upper() in row_text for k in self.empl_keys):
                    attendance_idx = idx
                    continue

            # detect payroll header
            if payroll_idx is None:
                if any(k.upper() in row_text for k in self.pay_keywords) and any(k.upper() in row_text for k in self.empl_keys):
                    payroll_idx = idx
                    continue

            # If we found both, break
            if attendance_idx is not None and payroll_idx is not None:
                break

        # Fallback heuristics: if one header not found try looser checks
        if attendance_idx is None:
            for idx, row in df.iterrows():
                row_values = [str(x) for x in row.values if not pd.isna(x)]
                if not row_values:
                    continue
                row_text = ' '.join(row_values).upper()
                if 'ATTENDANCE' in row_text and any(k.upper() in row_text for k in self.empl_keys):
                    attendance_idx = idx
                    break

        if payroll_idx is None:
            for idx, row in df.iterrows():
                row_values = [str(x) for x in row.values if not pd.isna(x)]
                if not row_values:
                    continue
                row_text = ' '.join(row_values).upper()
                if any(k.upper() in row_text for k in self.pay_keywords) and any(k.upper() in row_text for k in self.empl_keys):
                    payroll_idx = idx
                    break

        return attendance_idx, payroll_idx

    def _slice_table(self, df: pd.DataFrame, header_idx: int, next_header_idx: Optional[int]) -> pd.DataFrame:
        """Return a DataFrame slice for a table starting at header_idx (inclusive) until next_header_idx (exclusive)
        or end of df if next_header_idx is None."""
        if header_idx is None:
            return pd.DataFrame()
        if next_header_idx is not None and next_header_idx > header_idx:
            return df.iloc[header_idx:next_header_idx].reset_index(drop=True)
        else:
            return df.iloc[header_idx:].reset_index(drop=True)

    def _parse_table(self, table_df: pd.DataFrame, expected_type: str = 'attendance') -> Dict[str, Any]:
        """Parse a sliced table DataFrame. expected_type is 'attendance' or 'payroll'."""
        result = {
            'success': False,
            'file_name': None,
            'sheet_type': expected_type.title(),
            'employee_data': [],
            'total_rows': len(table_df),
            'total_columns': len(table_df.columns)
        }

        if table_df.empty:
            return result

        # First row is header
        header_row = table_df.iloc[0].fillna('').astype(str).tolist()
        norm_headers = [_normalize_header(h) for h in header_row]

        # Build mapping from normalized header -> column index
        header_map = {h: i for i, h in enumerate(norm_headers) if h}

        # iterate data rows
        data_rows = table_df.iloc[1:]
        employees = []

        for _, row in data_rows.iterrows():
            # stop at fully empty row
            if row.isna().all():
                continue
            row_vals = row.fillna('').tolist()
            # helper to get by fuzzy key
            def get_by_candidates(candidates: List[str]):
                for c in candidates:
                    key = _normalize_header(c)
                    if key in header_map:
                        return row_vals[header_map[key]]
                # try exact normalized keys present
                for hkey in header_map.keys():
                    for cand in candidates:
                        if cand.replace(' ', '') in hkey.replace(' ', ''):
                            return row_vals[header_map[hkey]]
                return ''

            empl_no = get_by_candidates(['empl no', 'empl.no', 's.no', 'empno', 'employee no'])
            emp_name = get_by_candidates(['employee name', 'employee', 'name'])

            if expected_type == 'attendance':
                attendance_days_raw = get_by_candidates(['attendance', 'attendance days', 'attendance/production types'])
                attendance_days = _parse_number(attendance_days_raw)
                emp = {
                    'empl_no': str(empl_no).strip(),
                    'employee_name': str(emp_name).strip(),
                    'attendance_days': attendance_days,
                    'raw_row': [str(x) for x in row_vals]
                }
            else:  # payroll
                # Pick common payroll columns
                basic = _parse_number(get_by_candidates(['basic']))
                hra = _parse_number(get_by_candidates(['hra']))
                medical = _parse_number(get_by_candidates(['medical']))
                gross = _parse_number(get_by_candidates(['gross', 'total', 'total gross', 'gross salary']))
                net = _parse_number(get_by_candidates(['net', 'net salary']))
                total = _parse_number(get_by_candidates(['total']))

                emp = {
                    'empl_no': str(empl_no).strip(),
                    'employee_name': str(emp_name).strip(),
                    'basic': basic,
                    'hra': hra,
                    'medical': medical,
                    'gross': gross,
                    'total': total,
                    'net': net,
                    'raw_row': [str(x) for x in row_vals]
                }

            employees.append(emp)

        result['success'] = True
        result['employee_data'] = employees
        result['total_employees'] = len(employees)
        return result

    def process_combined_sheet(self, file_path: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """Main entrypoint: process a single-sheet Excel file that contains both Attendance and Payroll tables.

        Returns a dict with keys 'attendance_result' and 'payroll_result'.
        """
        if not os.path.exists(file_path):
            return {'error': f'File not found: {file_path}', 'success': False}

        try:
            # Read first sheet by default
            excel = pd.ExcelFile(file_path)
            target_sheet = sheet_name or excel.sheet_names[0]
            df = pd.read_excel(file_path, sheet_name=target_sheet, header=None, dtype=object)

            attendance_idx, payroll_idx = self._find_header_indices(df)

            # If headers reversed or only one found, make best effort to assign order
            # Determine slices: sort indices
            indices = []
            if attendance_idx is not None:
                indices.append(('attendance', attendance_idx))
            if payroll_idx is not None:
                indices.append(('payroll', payroll_idx))

            # If none detected, try to find a single header row and heuristically split columns
            if not indices:
                # Try to find a single header row that contains 'EMPL NO'
                single_header = None
                for idx, row in df.iterrows():
                    row_values = [str(x) for x in row.values if not pd.isna(x)]
                    if not row_values:
                        continue
                    rt = ' '.join(row_values).upper()
                    if any(k.upper() in rt for k in self.empl_keys):
                        single_header = idx
                        break
                if single_header is None:
                    return {'error': 'Could not detect header rows for attendance or payroll', 'success': False}
                # In this case try to parse whole sheet as payroll and attendance separately (fall back)
                tbl = self._slice_table(df, single_header, None)
                payroll_result = self._parse_table(tbl, expected_type='payroll')
                attendance_result = self._parse_table(tbl, expected_type='attendance')
                return {
                    'attendance_result': attendance_result,
                    'payroll_result': payroll_result,
                    'success': True
                }

            # Sort indices by row number to slice properly
            indices_sorted = sorted(indices, key=lambda x: x[1])

            # Build slices between headers
            slices = {}
            for i, (typ, idx) in enumerate(indices_sorted):
                next_idx = indices_sorted[i + 1][1] if i + 1 < len(indices_sorted) else None
                slices[typ] = self._slice_table(df, idx, next_idx)

            attendance_tbl = slices.get('attendance', pd.DataFrame())
            payroll_tbl = slices.get('payroll', pd.DataFrame())

            attendance_result = self._parse_table(attendance_tbl, expected_type='attendance') if not attendance_tbl.empty else {'success': False}
            payroll_result = self._parse_table(payroll_tbl, expected_type='payroll') if not payroll_tbl.empty else {'success': False}

            # Attach some metadata
            attendance_result['file_name'] = os.path.basename(file_path)
            payroll_result['file_name'] = os.path.basename(file_path)

            return {
                'attendance_result': attendance_result,
                'payroll_result': payroll_result,
                'success': True
            }

        except Exception as e:
            return {'error': f'Error processing combined sheet: {str(e)}', 'success': False}


# Example quick CLI-style test helper (not executed by import)
if __name__ == '__main__':
    import sys
    p = CombinedProcessor()
    if len(sys.argv) < 2:
        print('Usage: python combined_processor.py <excel-file>')
        sys.exit(1)
    res = p.process_combined_sheet(sys.argv[1])
    import json
    print(json.dumps(res, indent=2, ensure_ascii=False))
