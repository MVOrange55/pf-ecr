"""
ECR Generator
=============
Validates uploaded Excel data and produces the EPFO ECR 2.0 text file.

ECR record format (#~# delimited, exactly 11 fields per employee line):
  UAN#~#NAME#~#GROSS#~#EPF_WAGES#~#EPS_WAGES#~#EDLI_WAGES#~#
  EE_SHARE#~#EPS_CONTRIB#~#ER_SHARE#~#NCP_DAYS#~#REFUND

Contribution rules:
  EPS Wages       = min(EPF Wages, 15 000)
  EDLI Wages      = min(EPF Wages, 15 000)
  EE Share        = round(EPF Wages × 12%)             [employee EPF]
  EPS Contrib     = round(EPS Wages × 8.33%)           [employer → pension]
  ER Share        = EE Share − EPS Contrib             [employer → EPF]
  EDLI Due        = round(EDLI Wages × 0.5%)           [informational only]
"""

from __future__ import annotations

import os
from typing import Any

import pandas as pd

# ── Constants ─────────────────────────────────────────────────────────────────
EPS_WAGE_CEILING    = 15_000
EPF_EMPLOYEE_RATE   = 0.12
EPS_RATE            = 0.0833
EDLI_RATE           = 0.005

# Accepted column aliases (all lower-cased)
COLUMN_ALIASES: dict[str, list[str]] = {
    'uan': [
        'uan', 'uan_number', 'uan no', 'uan number',
        'universal account number', 'universal_account_number',
    ],
    'member_name': [
        'member name', 'member_name', 'name', 'employee name',
        'emp name', 'employee_name', 'emp_name',
    ],
    'gross_wages': [
        'gross wages', 'gross_wages', 'gross salary', 'gross',
        'total wages', 'total_wages',
    ],
    'epf_wages': [
        'epf wages', 'epf_wages', 'basic', 'basic wages',
        'basic_wages', 'epf wage', 'basic + da', 'basic+da',
    ],
}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _find_col(df: pd.DataFrame, aliases: list[str]) -> str | None:
    for alias in aliases:
        if alias in df.columns:
            return alias
    return None


# ── Main validation & processing ──────────────────────────────────────────────

def validate_and_process(
    df: pd.DataFrame,
) -> tuple[pd.DataFrame | None, list[str], list[str]]:
    """
    Validate and enrich the raw DataFrame.

    Returns
    -------
    (processed_df | None, errors, warnings)
    If errors is non-empty, processed_df is None.
    """
    errors:   list[str] = []
    warnings: list[str] = []

    # Normalise column names
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip().str.lower()

    # Map to standard names
    col_map: dict[str, str] = {}
    for field, aliases in COLUMN_ALIASES.items():
        found = _find_col(df, aliases)
        if found is None:
            errors.append(
                f"Required column '{field}' not found. "
                f"Expected one of: {', '.join(aliases)}"
            )
        else:
            col_map[field] = found

    if errors:
        return None, errors, warnings

    # Rename to canonical names
    df = df.rename(columns={v: k for k, v in col_map.items()})

    # Optional columns
    for col, default in [('ncp_days', 0), ('refund_of_advances', 0)]:
        if col not in df.columns:
            df[col] = default

    # Drop fully-empty rows
    key_cols = ['uan', 'member_name', 'gross_wages', 'epf_wages']
    df = df.dropna(subset=key_cols, how='all').reset_index(drop=True)

    if df.empty:
        errors.append('No valid data rows found in the uploaded file.')
        return None, errors, warnings

    # ── Validate & clean UAN ─────────────────────────────────
    # Excel stores numeric UANs as float64 (e.g. 101988019020.0).
    # Convert via int() for floats to avoid the ".0" suffix and dtype issues.
    def _clean_uan(val) -> str:
        if pd.isna(val):
            return ''
        if isinstance(val, float):
            return str(int(val))          # 101988019020.0  →  '101988019020'
        return str(val).strip().rstrip('.0') if str(val).endswith('.0') else str(val).strip()

    df['uan'] = df['uan'].apply(_clean_uan)

    uan_errors: list[str] = []
    for idx, uan_str in df['uan'].items():
        uan_str = str(uan_str).strip()    # guarantee Python str regardless of pandas version
        row_num = int(idx) + 2            # spreadsheet row number (1-based + header)
        if not uan_str.isdigit():
            uan_errors.append(f'Row {row_num}: UAN "{uan_str}" must contain only digits.')
        elif len(uan_str) != 12:
            uan_errors.append(
                f'Row {row_num}: UAN "{uan_str}" must be exactly 12 digits '
                f'(found {len(uan_str)}).'
            )

    if uan_errors:
        errors.extend(uan_errors)
        return None, errors, warnings

    # ── Convert numeric columns ───────────────────────────────
    for col in ['gross_wages', 'epf_wages', 'ncp_days', 'refund_of_advances']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # ── Wage sanity checks ────────────────────────────────────
    for idx, row in df.iterrows():
        row_num = int(idx) + 2
        if row['gross_wages'] < 0:
            errors.append(f'Row {row_num}: Gross wages cannot be negative.')
        if row['epf_wages'] < 0:
            errors.append(f'Row {row_num}: EPF wages cannot be negative.')
        if row['epf_wages'] > row['gross_wages']:
            warnings.append(
                f'Row {row_num}: EPF wages ({row["epf_wages"]:,.0f}) '
                f'exceeds gross wages ({row["gross_wages"]:,.0f}).'
            )
        if not (0 <= row['ncp_days'] <= 31):
            warnings.append(f'Row {row_num}: NCP days ({row["ncp_days"]:.0f}) seems unusual.')

    if errors:
        return None, errors, warnings

    # ── Calculate contributions ────────────────────────────────
    df['eps_wages']  = df['epf_wages'].apply(lambda w: min(w, EPS_WAGE_CEILING))
    df['edli_wages'] = df['epf_wages'].apply(lambda w: min(w, EPS_WAGE_CEILING))

    df['epf_contrib_due']       = df['epf_wages'].apply(lambda w: round(w * EPF_EMPLOYEE_RATE))
    df['eps_contrib_due']       = df['eps_wages'].apply(lambda w: round(w * EPS_RATE))
    # Employer's share into EPF = total employer (12%) − EPS portion
    df['_employer_epf']         = df['epf_contrib_due'] - df['eps_contrib_due']
    df['epf_contrib_remitted']  = df['epf_contrib_due'] + df['_employer_epf']
    df['eps_contrib_remitted']  = df['eps_contrib_due']
    df['edli_contrib_due']      = df['edli_wages'].apply(lambda w: round(w * EDLI_RATE))
    df['edli_contrib_remitted'] = df['edli_contrib_due']

    # Informational warning about EPS cap
    capped = int((df['epf_wages'] > EPS_WAGE_CEILING).sum())
    if capped:
        warnings.append(
            f'{capped} employee(s) have EPF wages above ₹{EPS_WAGE_CEILING:,}. '
            f'EPS wages have been capped at ₹{EPS_WAGE_CEILING:,}.'
        )

    return df, errors, warnings


# ── ECR file writer ────────────────────────────────────────────────────────────

def generate_ecr_file(
    df: pd.DataFrame,
    company: Any,
    wage_month: str,    # "MM/YYYY"
    ecr_folder: str,
) -> tuple[str, str, dict]:
    """
    Write the ECR text file and return (filepath, filename, summary_dict).
    """
    month, year = wage_month.split('/')

    lines: list[str] = []
    for _, row in df.iterrows():
        uan  = str(row['uan']).strip()
        name = str(row['member_name']).strip()
        if not uan:          # skip blank rows that slipped through
            continue
        # EPFO ECR 2.0: exactly 11 fields separated by #~#
        # (same layout as successful challan uploads, e.g. GURJAR_ECR_CHAL)
        # ER share = employer EPF portion (EE 12% − EPS 8.33%)
        er_share = int(row['epf_contrib_due']) - int(row['eps_contrib_due'])
        fields = [
            uan,
            name,
            str(int(row['gross_wages'])),
            str(int(row['epf_wages'])),
            str(int(row['eps_wages'])),
            str(int(row['edli_wages'])),
            str(int(row['epf_contrib_due'])),
            str(int(row['eps_contrib_due'])),
            str(er_share),
            str(int(row['ncp_days'])),
            str(int(row['refund_of_advances'])),
        ]
        lines.append('#~#'.join(fields))

    # CRLF line endings; no blank lines; no header row
    content = '\r\n'.join(line for line in lines if line.strip())

    # Build filename: CompanyName_MMYYYY_ECR.txt
    safe_name = ''.join(c if (c.isalnum() or c == '_') else '_' for c in company.name)
    filename  = f'{safe_name}_{month}{year}_ECR.txt'

    os.makedirs(ecr_folder, exist_ok=True)
    filepath = os.path.join(ecr_folder, filename)

    # Avoid overwriting previous files
    base, ext = os.path.splitext(filepath)
    counter   = 1
    while os.path.exists(filepath):
        filepath = f'{base}_{counter}{ext}'
        filename = os.path.basename(filepath)
        counter += 1

    # newline='' so Windows does not turn \r\n into \r\r\n
    with open(filepath, 'w', encoding='utf-8', newline='') as fh:
        fh.write(content)

    summary = {
        'employee_count':        len(df),
        'total_gross_wages':     float(df['gross_wages'].sum()),
        'total_epf_wages':       float(df['epf_wages'].sum()),
        'total_eps_wages':       float(df['eps_wages'].sum()),
        'total_epf_contribution': float(df['epf_contrib_remitted'].sum()),
        'total_eps_contribution': float(df['eps_contrib_remitted'].sum()),
        'total_edli_contribution': float(df['edli_contrib_remitted'].sum()),
    }
    return filepath, filename, summary


# ── ECR → Excel (reverse) ──────────────────────────────────────────────────────

ECR_XLS_COLUMNS = [
    'UAN', 'Member Name', 'Gross Wages', 'EPF Wages', 'EPS Wages',
    'EDLI Wages', 'EE Share', 'EPS Share', 'ER Share', 'NCP Days',
    'Refund of Advances',
]


def parse_ecr_lines(lines) -> pd.DataFrame:
    """
    Parse EPFO ECR text lines into a DataFrame.

    Supports current `#~#` delimiter and legacy `#`-delimited lines.
    """
    rows: list[list[str]] = []
    for raw in lines:
        line = raw.strip() if isinstance(raw, str) else str(raw).strip()
        if not line or line in ('#~#', '~'):
            continue

        if '#~#' in line:
            parts = [p.strip() for p in line.split('#~#')]
        else:
            # Legacy: #UAN#NAME#...#  or  UAN#NAME#...
            parts = [p.strip() for p in line.split('#') if p.strip() != '']

        if len(parts) < 4:
            continue

        # Pad / trim to 11 fields
        if len(parts) < 11:
            parts = parts + ['0'] * (11 - len(parts))
        rows.append(parts[:11])

    if not rows:
        raise ValueError('No valid ECR records found in the file.')

    df = pd.DataFrame(rows, columns=ECR_XLS_COLUMNS)
    for col in ECR_XLS_COLUMNS:
        if col in ('UAN', 'Member Name'):
            continue
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    df['UAN'] = df['UAN'].astype(str).str.strip()
    return df


def parse_ecr_txt(filepath: str) -> pd.DataFrame:
    """Parse an EPFO ECR .txt file into a DataFrame."""
    with open(filepath, encoding='utf-8-sig', newline='') as fh:
        return parse_ecr_lines(fh)


def parse_ecr_bytes(data: bytes) -> pd.DataFrame:
    """Parse ECR content from uploaded file bytes."""
    text = data.decode('utf-8-sig', errors='replace')
    return parse_ecr_lines(text.splitlines())


def _dataframe_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    from io import BytesIO

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='ECR')
    return buf.getvalue()


def ecr_file_to_xlsx_bytes(filepath: str) -> bytes:
    """Convert an ECR .txt file to an .xlsx workbook (bytes)."""
    return _dataframe_to_xlsx_bytes(parse_ecr_txt(filepath))


def ecr_bytes_to_xlsx_bytes(data: bytes) -> bytes:
    """Convert uploaded ECR file bytes to an .xlsx workbook (bytes)."""
    return _dataframe_to_xlsx_bytes(parse_ecr_bytes(data))
