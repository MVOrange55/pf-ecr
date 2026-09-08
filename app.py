from __future__ import annotations

import os
from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st


# ============================================================
# CONSTANTS
# ============================================================

EPS_WAGE_CEILING = 15_000
EPF_EMPLOYEE_RATE = 0.12
EPS_RATE = 0.0833
EDLI_RATE = 0.005

ECR_XLS_COLUMNS = [
    "UAN",
    "Member Name",
    "Gross Wages",
    "EPF Wages",
    "EPS Wages",
    "EDLI Wages",
    "EE Share",
    "EPS Share",
    "ER Share",
    "NCP Days",
    "Refund of Advances",
]

COLUMN_ALIASES: dict[str, list[str]] = {
    "uan": [
        "uan",
        "uan_number",
        "uan no",
        "uan number",
        "universal account number",
        "universal_account_number",
    ],
    "member_name": [
        "member name",
        "member_name",
        "name",
        "employee name",
        "emp name",
        "employee_name",
        "emp_name",
    ],
    "gross_wages": [
        "gross wages",
        "gross_wages",
        "gross salary",
        "gross",
        "total wages",
        "total_wages",
    ],
    "epf_wages": [
        "epf wages",
        "epf_wages",
        "basic",
        "basic wages",
        "basic_wages",
        "epf wage",
        "basic + da",
        "basic+da",
    ],
}


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def _normalise_column_name(value: Any) -> str:
    """Normalise Excel column names."""
    return (
        str(value)
        .strip()
        .lower()
        .replace("\n", " ")
    )


def _find_col(df: pd.DataFrame, aliases: list[str]) -> str | None:
    """Find the first matching column."""
    for alias in aliases:
        if alias in df.columns:
            return alias
    return None


def _clean_uan(value: Any) -> str:
    """
    Safely convert Excel UAN values to a 12-digit string.

    Handles:
        101988019020
        101988019020.0
        "101988019020"
    """
    if pd.isna(value):
        return ""

    # Excel may read large numeric UAN as float.
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).strip()

    value = str(value).strip()

    # Remove ONLY literal trailing ".0"
    if value.endswith(".0"):
        value = value[:-2]

    return value


def _safe_int(value: Any, default: int = 0) -> int:
    """Convert a value safely to integer."""
    try:
        if pd.isna(value):
            return default
        return int(float(value))
    except (ValueError, TypeError):
        return default


def _format_amount(value: Any) -> str:
    """Format contribution/wage as integer text."""
    return str(_safe_int(value))


# ============================================================
# VALIDATION AND PROCESSING
# ============================================================

def validate_and_process(
    df: pd.DataFrame,
) -> tuple[pd.DataFrame | None, list[str], list[str]]:
    """
    Validate and enrich uploaded Excel DataFrame.

    Returns:
        processed_df, errors, warnings
    """

    errors: list[str] = []
    warnings: list[str] = []

    if df is None or df.empty:
        return None, ["Uploaded Excel file is empty."], warnings

    df = df.copy()

    # --------------------------------------------------------
    # Preserve original Excel row numbers
    # --------------------------------------------------------

    # Header is row 1, so first data row is row 2.
    df["_excel_row"] = range(2, len(df) + 2)

    # --------------------------------------------------------
    # Normalise columns
    # --------------------------------------------------------

    df.columns = [_normalise_column_name(c) for c in df.columns]

    # --------------------------------------------------------
    # Find required columns
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # Rename columns to canonical names
    # --------------------------------------------------------

    rename_map = {
        source: target
        for target, source in col_map.items()
    }

    df = df.rename(columns=rename_map)

    # --------------------------------------------------------
    # Optional columns
    # --------------------------------------------------------

    optional_columns = {
        "ncp_days": 0,
        "refund_of_advances": 0,
    }

    for col, default in optional_columns.items():
        if col not in df.columns:
            df[col] = default

    # --------------------------------------------------------
    # Remove completely empty rows
    # --------------------------------------------------------

    key_cols = [
        "uan",
        "member_name",
        "gross_wages",
        "epf_wages",
    ]

    df = df.dropna(
        subset=key_cols,
        how="all",
    ).reset_index(drop=True)

    if df.empty:
        return None, ["No valid data rows found in the uploaded file."], warnings

    # --------------------------------------------------------
    # Clean UAN
    # --------------------------------------------------------

    df["uan"] = df["uan"].apply(_clean_uan)

    uan_errors: list[str] = []

    for _, row in df.iterrows():

        uan_str = str(row["uan"]).strip()
        row_num = int(row["_excel_row"])

        if not uan_str:
            uan_errors.append(
                f"Row {row_num}: UAN is blank."
            )
            continue

        if not uan_str.isdigit():
            uan_errors.append(
                f'Row {row_num}: UAN "{uan_str}" '
                f"must contain only digits."
            )
            continue

        if len(uan_str) != 12:
            uan_errors.append(
                f'Row {row_num}: UAN "{uan_str}" '
                f"must be exactly 12 digits "
                f"(found {len(uan_str)})."
            )

    if uan_errors:
        errors.extend(uan_errors)
        return None, errors, warnings

    # --------------------------------------------------------
    # Convert numeric columns
    # --------------------------------------------------------

    numeric_columns = [
        "gross_wages",
        "epf_wages",
        "ncp_days",
        "refund_of_advances",
    ]

    for col in numeric_columns:

        original = df[col].copy()

        converted = pd.to_numeric(
            original,
            errors="coerce",
        )

        # Identify non-empty values that could not convert.
        for idx, value in original.items():

            if pd.isna(value):
                continue

            converted_value = converted.loc[idx]

            if pd.isna(converted_value):
                row_num = int(df.loc[idx, "_excel_row"])

                errors.append(
                    f"Row {row_num}: "
                    f"Invalid numeric value in '{col}': "
                    f'"{value}".'
                )

        df[col] = converted.fillna(0)

    if errors:
        return None, errors, warnings

    # --------------------------------------------------------
    # Wage validation
    # --------------------------------------------------------

    for idx, row in df.iterrows():

        row_num = int(row["_excel_row"])

        gross = float(row["gross_wages"])
        epf = float(row["epf_wages"])
        ncp = float(row["ncp_days"])
        refund = float(row["refund_of_advances"])

        if gross < 0:
            errors.append(
                f"Row {row_num}: Gross wages cannot be negative."
            )

        if epf < 0:
            errors.append(
                f"Row {row_num}: EPF wages cannot be negative."
            )

        if refund < 0:
            warnings.append(
                f"Row {row_num}: Refund of Advances is negative."
            )

        if epf > gross:
            warnings.append(
                f"Row {row_num}: EPF wages "
                f"({epf:,.0f}) exceeds gross wages "
                f"({gross:,.0f})."
            )

        if not 0 <= ncp <= 31:
            warnings.append(
                f"Row {row_num}: NCP days "
                f"({ncp:.0f}) seems unusual."
            )

    if errors:
        return None, errors, warnings

    # --------------------------------------------------------
    # Calculate ECR wages
    # --------------------------------------------------------

    df["eps_wages"] = df["epf_wages"].apply(
        lambda wage: min(wage, EPS_WAGE_CEILING)
    )

    df["edli_wages"] = df["epf_wages"].apply(
        lambda wage: min(wage, EPS_WAGE_CEILING)
    )

    # --------------------------------------------------------
    # Calculate contributions
    # --------------------------------------------------------

    # Employee EPF contribution = 12% of EPF wages.
    df["ee_share"] = df["epf_wages"].apply(
        lambda wage: round(wage * EPF_EMPLOYEE_RATE)
    )

    # Employer pension contribution = 8.33% of capped EPS wages.
    df["eps_share"] = df["eps_wages"].apply(
        lambda wage: round(wage * EPS_RATE)
    )

    # Employer EPF share.
    df["er_share"] = (
        df["ee_share"] - df["eps_share"]
    )

    # EDLI is informational and is not one of the 11 ECR fields.
    df["edli_contrib_due"] = df["edli_wages"].apply(
        lambda wage: round(wage * EDLI_RATE)
    )

    # --------------------------------------------------------
    # EPS cap warning
    # --------------------------------------------------------

    capped_count = int(
        (df["epf_wages"] > EPS_WAGE_CEILING).sum()
    )

    if capped_count:
        warnings.append(
            f"{capped_count} employee(s) have EPF wages above "
            f"₹{EPS_WAGE_CEILING:,}. "
            f"EPS and EDLI wages have been capped at "
            f"₹{EPS_WAGE_CEILING:,}."
        )

    # --------------------------------------------------------
    # Compatibility columns
    # --------------------------------------------------------

    # Keep these names if other parts of the application use them.
    df["epf_contrib_due"] = df["ee_share"]
    df["eps_contrib_due"] = df["eps_share"]

    # Total EPF remitted = employee share + employer EPF share.
    df["epf_contrib_remitted"] = (
        df["ee_share"] + df["er_share"]
    )

    df["eps_contrib_remitted"] = df["eps_share"]
    df["edli_contrib_remitted"] = df["edli_contrib_due"]

    return df, errors, warnings


# ============================================================
# ECR FILE GENERATOR
# ============================================================

def generate_ecr_file(
    df: pd.DataFrame,
    company: Any,
    wage_month: str,
    ecr_folder: str,
) -> tuple[str, str, dict]:
    """
    Generate EPFO ECR 2.0 text file.

    Format:

    UAN#~#NAME#~#GROSS#~#EPF_WAGES#~#EPS_WAGES#~#
    EDLI_WAGES#~#EE_SHARE#~#EPS_SHARE#~#ER_SHARE#~#
    NCP_DAYS#~#REFUND
    """

    if "/" not in wage_month:
        raise ValueError(
            "Wage month must be in MM/YYYY format."
        )

    month, year = wage_month.split("/")

    if len(month) != 2 or len(year) != 4:
        raise ValueError(
            "Wage month must be in MM/YYYY format."
        )

    month_int = int(month)

    if not 1 <= month_int <= 12:
        raise ValueError("Invalid month.")

    lines: list[str] = []

    for _, row in df.iterrows():

        uan = str(row["uan"]).strip()
        name = str(row["member_name"]).strip()

        if not uan:
            continue

        ee_share = _safe_int(row["ee_share"])
        eps_share = _safe_int(row["eps_share"])

        # Employer EPF portion.
        er_share = ee_share - eps_share

        fields = [
            uan,
            name,
            _format_amount(row["gross_wages"]),
            _format_amount(row["epf_wages"]),
            _format_amount(row["eps_wages"]),
            _format_amount(row["edli_wages"]),
            str(ee_share),
            str(eps_share),
            str(er_share),
            str(_safe_int(row["ncp_days"])),
            str(_safe_int(row["refund_of_advances"])),
        ]

        # Exactly 11 fields.
        if len(fields) != 11:
            raise ValueError(
                f"Internal error: ECR record contains "
                f"{len(fields)} fields instead of 11."
            )

        lines.append("#~#".join(fields))

    if not lines:
        raise ValueError(
            "No valid employee records available for ECR."
        )

    # CRLF line endings, no header.
    content = "\r\n".join(lines)

    # --------------------------------------------------------
    # Company name
    # --------------------------------------------------------

    company_name = getattr(
        company,
        "name",
        str(company),
    )

    company_name = str(company_name).strip()

    safe_name = "".join(
        c if (c.isalnum() or c == "_") else "_"
        for c in company_name
    )

    safe_name = safe_name.strip("_") or "Company"

    filename = (
        f"{safe_name}_{month}{year}_ECR.txt"
    )

    os.makedirs(
        ecr_folder,
        exist_ok=True,
    )

    filepath = os.path.join(
        ecr_folder,
        filename,
    )

    # Avoid overwriting.
    base, ext = os.path.splitext(filepath)

    counter = 1

    while os.path.exists(filepath):

        filepath = (
            f"{base}_{counter}{ext}"
        )

        filename = os.path.basename(filepath)

        counter += 1

    with open(
        filepath,
        "w",
        encoding="utf-8",
        newline="",
    ) as fh:
        fh.write(content)

    summary = {
        "employee_count": len(df),
        "total_gross_wages": float(
            df["gross_wages"].sum()
        ),
        "total_epf_wages": float(
            df["epf_wages"].sum()
        ),
        "total_eps_wages": float(
            df["eps_wages"].sum()
        ),
        "total_edli_wages": float(
            df["edli_wages"].sum()
        ),
        "total_ee_share": float(
            df["ee_share"].sum()
        ),
        "total_eps_contribution": float(
            df["eps_share"].sum()
        ),
        "total_er_share": float(
            df["er_share"].sum()
        ),
        "total_edli_contribution": float(
            df["edli_contrib_due"].sum()
        ),
    }

    return filepath, filename, summary


# ============================================================
# ECR PARSER
# ============================================================

def parse_ecr_lines(lines) -> pd.DataFrame:
    """
    Parse ECR text lines.

    Supports:
        #~# delimiter
        Legacy # delimiter
    """

    rows: list[list[str]] = []

    for raw in lines:

        line = (
            raw.strip()
            if isinstance(raw, str)
            else str(raw).strip()
        )

        if not line:
            continue

        if line in ("#~#", "~"):
            continue

        # ----------------------------------------------------
        # Current ECR format
        # ----------------------------------------------------

        if "#~#" in line:

            parts = [
                p.strip()
                for p in line.split("#~#")
            ]

        # ----------------------------------------------------
        # Legacy format
        # ----------------------------------------------------

        else:

            parts = [
                p.strip()
                for p in line.split("#")
                if p.strip() != ""
            ]

        # ----------------------------------------------------
        # Skip invalid lines
        # ----------------------------------------------------

        if len(parts) < 4:
            continue

        # Pad to 11 fields.
        if len(parts) < 11:
            parts.extend(
                ["0"] * (11 - len(parts))
            )

        # Trim anything beyond 11.
        rows.append(parts[:11])

    if not rows:
        raise ValueError(
            "No valid ECR records found in the file."
        )

    df = pd.DataFrame(
        rows,
        columns=ECR_XLS_COLUMNS,
    )

    # --------------------------------------------------------
    # Clean text columns
    # --------------------------------------------------------

    df["UAN"] = (
        df["UAN"]
        .astype(str)
        .str.strip()
    )

    df["Member Name"] = (
        df["Member Name"]
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # Numeric columns
    # --------------------------------------------------------

    for col in ECR_XLS_COLUMNS:

        if col in ("UAN", "Member Name"):
            continue

        df[col] = pd.to_numeric(
            df[col],
            errors="coerce",
        ).fillna(0).astype(int)

    return df


def parse_ecr_txt(filepath: str) -> pd.DataFrame:
    """Parse an ECR TXT file."""

    with open(
        filepath,
        encoding="utf-8-sig",
        newline="",
    ) as fh:

        return parse_ecr_lines(fh)


def parse_ecr_bytes(data: bytes) -> pd.DataFrame:
    """Parse uploaded ECR bytes."""

    text = data.decode(
        "utf-8-sig",
        errors="replace",
    )

    return parse_ecr_lines(
        text.splitlines()
    )


# ============================================================
# EXCEL CONVERSION
# ============================================================

def dataframe_to_xlsx_bytes(
    df: pd.DataFrame,
) -> bytes:

    buf = BytesIO()

    with pd.ExcelWriter(
        buf,
        engine="openpyxl",
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="ECR",
        )

    return buf.getvalue()


def ecr_file_to_xlsx_bytes(
    filepath: str,
) -> bytes:

    return dataframe_to_xlsx_bytes(
        parse_ecr_txt(filepath)
    )


def ecr_bytes_to_xlsx_bytes(
    data: bytes,
) -> bytes:

    return dataframe_to_xlsx_bytes(
        parse_ecr_bytes(data)
    )


# ============================================================
# STREAMLIT UI
# ============================================================

st.set_page_config(
    page_title="ECR Generator",
    page_icon="📄",
    layout="wide",
)


st.title("📄 EPFO ECR 2.0 Generator")

st.caption(
    "Validate Excel employee data, calculate contributions, "
    "generate ECR TXT, and convert ECR TXT back to Excel."
)


# ============================================================
# TABS
# ============================================================

tab_generate, tab_convert = st.tabs(
    [
        "📊 Excel → ECR",
        "📄 ECR → Excel",
    ]
)


# ============================================================
# TAB 1: EXCEL → ECR
# ============================================================

with tab_generate:

    st.subheader("Upload Employee Excel")

    uploaded_excel = st.file_uploader(
        "Choose Excel file",
        type=["xlsx", "xls"],
        key="employee_excel",
    )

    company_name = st.text_input(
        "Company Name",
        placeholder="Example: ABC PRIVATE LIMITED",
    )

    wage_month = st.text_input(
        "Wage Month",
        value="08/2026",
        help="Enter month in MM/YYYY format.",
    )

    if uploaded_excel is not None:

        try:

            df_input = pd.read_excel(
                uploaded_excel,
                engine="openpyxl",
            )

            st.success(
                f"Excel loaded successfully: "
                f"{len(df_input)} rows."
            )

            with st.expander(
                "Preview uploaded data"
            ):

                st.dataframe(
                    df_input,
                    use_container_width=True,
                )

        except Exception as exc:

            st.error(
                f"Could not read Excel file: {exc}"
            )

            df_input = None

    else:
        df_input = None

    # --------------------------------------------------------
    # Generate button
    # --------------------------------------------------------

    if st.button(
        "🚀 Validate & Generate ECR",
        type="primary",
        use_container_width=True,
    ):

        if df_input is None:
            st.error(
                "Please upload an Excel file first."
            )
            st.stop()

        if not company_name.strip():
            st.error(
                "Please enter Company Name."
            )
            st.stop()

        # Validate month before processing.
        try:

            parts = wage_month.strip().split("/")

            if len(parts) != 2:
                raise ValueError

            month = int(parts[0])
            year = int(parts[1])

            if not (
                1 <= month <= 12
                and 1900 <= year <= 2100
            ):
                raise ValueError

            wage_month_clean = (
                f"{month:02d}/{year:04d}"
            )

        except ValueError:

            st.error(
                "Invalid Wage Month. "
                "Use MM/YYYY, for example 08/2026."
            )

            st.stop()

        # ----------------------------------------------------
        # Validate
        # ----------------------------------------------------

        processed_df, errors, warnings = (
            validate_and_process(df_input)
        )

        # ----------------------------------------------------
        # Errors
        # ----------------------------------------------------

        if errors:

            st.error(
                f"Validation failed with "
                f"{len(errors)} error(s)."
            )

            for error in errors:
                st.error(error)

            if warnings:

                st.warning("Warnings:")

                for warning in warnings:
                    st.warning(warning)

            st.stop()

        # ----------------------------------------------------
        # Warnings
        # ----------------------------------------------------

        if warnings:

            st.warning(
                f"{len(warnings)} warning(s) found."
            )

            with st.expander(
                "View warnings"
            ):

                for warning in warnings:
                    st.write(f"⚠️ {warning}")

        # ----------------------------------------------------
        # Show processed data
        # ----------------------------------------------------

        st.subheader(
            "Calculated ECR Data"
        )

        display_columns = [
            "uan",
            "member_name",
            "gross_wages",
            "epf_wages",
            "eps_wages",
            "edli_wages",
            "ee_share",
            "eps_share",
            "er_share",
            "ncp_days",
            "refund_of_advances",
        ]

        st.dataframe(
            processed_df[display_columns],
            use_container_width=True,
        )

        # ----------------------------------------------------
        # Summary
        # ----------------------------------------------------

        st.subheader("Summary")

        col1, col2, col3, col4 = st.columns(4)

        col1.metric(
            "Employees",
            len(processed_df),
        )

        col2.metric(
            "Gross Wages",
            f"₹{processed_df['gross_wages'].sum():,.0f}",
        )

        col3.metric(
            "EPF Wages",
            f"₹{processed_df['epf_wages'].sum():,.0f}",
        )

        col4.metric(
            "EPS Wages",
            f"₹{processed_df['eps_wages'].sum():,.0f}",
        )

        col5, col6, col7, col8 = st.columns(4)

        col5.metric(
            "EE Share",
            f"₹{processed_df['ee_share'].sum():,.0f}",
        )

        col6.metric(
            "EPS Share",
            f"₹{processed_df['eps_share'].sum():,.0f}",
        )

        col7.metric(
            "ER Share",
            f"₹{processed_df['er_share'].sum():,.0f}",
        )

        col8.metric(
            "EDLI Due",
            f"₹{processed_df['edli_contrib_due'].sum():,.0f}",
        )

        # ----------------------------------------------------
        # Generate ECR in memory
        # ----------------------------------------------------

        try:

            class Company:
                def __init__(self, name):
                    self.name = name

            company = Company(
                company_name.strip()
            )

            # Use temporary application folder.
            ecr_folder = os.path.join(
                os.getcwd(),
                "ecr_output",
            )

            filepath, filename, summary = (
                generate_ecr_file(
                    processed_df,
                    company,
                    wage_month_clean,
                    ecr_folder,
                )
            )

            with open(
                filepath,
                "rb",
            ) as fh:

                ecr_bytes = fh.read()

            st.success(
                f"ECR generated successfully: {filename}"
            )

            # ------------------------------------------------
            # Download
            # ------------------------------------------------

            st.download_button(
                label="⬇️ Download ECR TXT",
                data=ecr_bytes,
                file_name=filename,
                mime="text/plain",
                type="primary",
                use_container_width=True,
            )

        except Exception as exc:

            st.error(
                f"Could not generate ECR file: {exc}"
            )


# ============================================================
# TAB 2: ECR → EXCEL
# ============================================================

with tab_convert:

    st.subheader(
        "Convert ECR TXT to Excel"
    )

    uploaded_ecr = st.file_uploader(
        "Choose ECR TXT file",
        type=["txt"],
        key="ecr_txt",
    )

    if uploaded_ecr is not None:

        try:

            ecr_data = uploaded_ecr.read()

            converted_df = parse_ecr_bytes(
                ecr_data
            )

            st.success(
                f"ECR parsed successfully: "
                f"{len(converted_df)} records."
            )

            st.dataframe(
                converted_df,
                use_container_width=True,
            )

            xlsx_bytes = (
                ecr_bytes_to_xlsx_bytes(
                    ecr_data
                )
            )

            output_filename = (
                os.path.splitext(
                    uploaded_ecr.name
                )[0]
                + ".xlsx"
            )

            st.download_button(
                label="⬇️ Download Excel",
                data=xlsx_bytes,
                file_name=output_filename,
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                type="primary",
                use_container_width=True,
            )

        except Exception as exc:

            st.error(
                f"Could not parse ECR file: {exc}"
            )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "ECR Generator • EPFO ECR 2.0 • "
    "11-field #~# delimited format"
)
