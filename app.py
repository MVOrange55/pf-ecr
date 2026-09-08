from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="EPFO ECR Generator",
    page_icon="📄",
    layout="wide",
)


# ============================================================
# CONSTANTS
# ============================================================

EPS_WAGE_CEILING = 15000

EPF_RATE = 0.12
EPS_RATE = 0.0833
EDLI_RATE = 0.005


# ============================================================
# SESSION STATE
# ============================================================

if "employees" not in st.session_state:
    st.session_state.employees = []


# ============================================================
# HELPERS
# ============================================================

def clean_number(value: Any) -> int:
    """Convert a value into a non-negative integer."""
    try:
        return max(0, int(float(value)))
    except (ValueError, TypeError):
        return 0


def calculate_employee(
    uan: str,
    name: str,
    gross: float,
    epf_wages: float,
    ncp_days: int,
    refund: float,
):
    """Calculate contribution values for one employee."""

    gross = max(0, float(gross))
    epf_wages = max(0, float(epf_wages))
    ncp_days = max(0, int(ncp_days))
    refund = max(0, float(refund))

    eps_wages = min(epf_wages, EPS_WAGE_CEILING)
    edli_wages = min(epf_wages, EPS_WAGE_CEILING)

    # Employee EPF share = 12%
    ee_share = int(round(epf_wages * EPF_RATE))

    # Employer pension share = 8.33%
    eps_share = int(round(eps_wages * EPS_RATE))

    # Employer EPF share
    er_share = max(0, ee_share - eps_share)

    # EDLI - informational
    edli_due = int(round(edli_wages * EDLI_RATE))

    return {
        "UAN": str(uan).strip(),
        "Member Name": str(name).strip(),
        "Gross Wages": clean_number(gross),
        "EPF Wages": clean_number(epf_wages),
        "EPS Wages": clean_number(eps_wages),
        "EDLI Wages": clean_number(edli_wages),
        "EE Share": clean_number(ee_share),
        "EPS Share": clean_number(eps_share),
        "ER Share": clean_number(er_share),
        "NCP Days": clean_number(ncp_days),
        "Refund of Advances": clean_number(refund),
        "EDLI Due": clean_number(edli_due),
    }


def validate_employee(
    uan,
    name,
    gross,
    epf_wages,
    ncp_days,
    refund,
):
    """Validate employee data."""

    errors = []
    warnings = []

    uan = str(uan).strip()
    name = str(name).strip()

    # --------------------------------------------------------
    # UAN
    # --------------------------------------------------------

    if not uan:
        errors.append("UAN is required.")

    elif not uan.isdigit():
        errors.append("UAN must contain digits only.")

    elif len(uan) != 12:
        errors.append(
            f"UAN must be exactly 12 digits. Currently {len(uan)} digits."
        )

    # --------------------------------------------------------
    # Name
    # --------------------------------------------------------

    if not name:
        errors.append("Member name is required.")

    # --------------------------------------------------------
    # Gross wages
    # --------------------------------------------------------

    try:
        gross = float(gross)
    except (ValueError, TypeError):
        errors.append("Gross wages must be a number.")
        gross = 0

    if gross < 0:
        errors.append("Gross wages cannot be negative.")

    # --------------------------------------------------------
    # EPF wages
    # --------------------------------------------------------

    try:
        epf_wages = float(epf_wages)
    except (ValueError, TypeError):
        errors.append("EPF wages must be a number.")
        epf_wages = 0

    if epf_wages < 0:
        errors.append("EPF wages cannot be negative.")

    if epf_wages > gross:
        warnings.append(
            "EPF wages are greater than gross wages. Please verify the wages."
        )

    # --------------------------------------------------------
    # NCP days
    # --------------------------------------------------------

    try:
        ncp_days = int(float(ncp_days))
    except (ValueError, TypeError):
        errors.append("NCP days must be a number.")
        ncp_days = 0

    if ncp_days < 0:
        errors.append("NCP days cannot be negative.")

    if ncp_days > 31:
        errors.append("NCP days cannot be greater than 31.")

    # --------------------------------------------------------
    # Refund
    # --------------------------------------------------------

    try:
        refund = float(refund)
    except (ValueError, TypeError):
        errors.append("Refund of Advances must be a number.")
        refund = 0

    if refund < 0:
        errors.append("Refund of Advances cannot be negative.")

    return errors, warnings


def employee_to_ecr_line(employee: dict) -> str:
    """
    Convert employee data to a pipe-separated ECR-style line.

    NOTE:
    Verify the exact field order/format against the current EPFO
    ECR specification before uploading to the EPFO portal.
    """

    fields = [
        employee["UAN"],
        employee["Member Name"],
        employee["Gross Wages"],
        employee["EPF Wages"],
        employee["EPS Wages"],
        employee["EDLI Wages"],
        employee["EE Share"],
        employee["EPS Share"],
        employee["ER Share"],
        employee["NCP Days"],
        employee["Refund of Advances"],
    ]

    return "|".join(str(x) for x in fields)


def generate_ecr_text(employees: list[dict]) -> str:
    """Generate ECR text."""

    lines = []

    for employee in employees:
        lines.append(employee_to_ecr_line(employee))

    return "\n".join(lines)


def generate_excel(employees: list[dict]) -> bytes:
    """Generate Excel file."""

    df = pd.DataFrame(employees)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="ECR Data",
        )

    return output.getvalue()


def reset_form():
    """Clear form values."""

    st.session_state.uan = ""
    st.session_state.member_name = ""
    st.session_state.gross_wages = 0
    st.session_state.epf_wages = 0
    st.session_state.ncp_days = 0
    st.session_state.refund = 0


# ============================================================
# HEADER
# ============================================================

st.title("📄 EPFO ECR Generator")

st.caption(
    "Employee contribution calculator, validation and ECR-style text export"
)

st.divider()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Settings")

    st.info(
        "Contribution rates used by this application:"
    )

    st.write(f"**EPF Rate:** {EPF_RATE * 100:.2f}%")
    st.write(f"**EPS Rate:** {EPS_RATE * 100:.2f}%")
    st.write(f"**EDLI Rate:** {EDLI_RATE * 100:.2f}%")
    st.write(f"**EPS Wage Ceiling:** ₹{EPS_WAGE_CEILING:,}")

    st.divider()

    if st.button(
        "🗑️ Clear All Employees",
        use_container_width=True,
    ):
        st.session_state.employees = []
        st.rerun()


# ============================================================
# EMPLOYEE FORM
# ============================================================

st.header("➕ Add Employee")

with st.form("employee_form"):

    col1, col2, col3 = st.columns(3)

    with col1:

        uan = st.text_input(
            "UAN *",
            key="uan",
            max_chars=12,
            placeholder="12 digit UAN",
        )

        member_name = st.text_input(
            "Member Name *",
            key="member_name",
            placeholder="Employee full name",
        )

    with col2:

        gross_wages = st.number_input(
            "Gross Wages",
            min_value=0,
            value=0,
            step=100,
            key="gross_wages",
        )

        epf_wages = st.number_input(
            "EPF Wages",
            min_value=0,
            value=0,
            step=100,
            key="epf_wages",
        )

    with col3:

        ncp_days = st.number_input(
            "NCP Days",
            min_value=0,
            max_value=31,
            value=0,
            step=1,
            key="ncp_days",
        )

        refund = st.number_input(
            "Refund of Advances",
            min_value=0,
            value=0,
            step=100,
            key="refund",
        )

    submitted = st.form_submit_button(
        "➕ Add Employee",
        use_container_width=True,
        type="primary",
    )


# ============================================================
# ADD EMPLOYEE
# ============================================================

if submitted:

    errors, warnings = validate_employee(
        uan=uan,
        name=member_name,
        gross=gross_wages,
        epf_wages=epf_wages,
        ncp_days=ncp_days,
        refund=refund,
    )

    if errors:

        for error in errors:
            st.error(error)

    else:

        for warning in warnings:
            st.warning(warning)

        # Check duplicate UAN
        duplicate = any(
            employee["UAN"] == str(uan).strip()
            for employee in st.session_state.employees
        )

        if duplicate:

            st.error(
                f"UAN {uan} already exists in the employee list."
            )

        else:

            employee = calculate_employee(
                uan=uan,
                name=member_name,
                gross=gross_wages,
                epf_wages=epf_wages,
                ncp_days=ncp_days,
                refund=refund,
            )

            st.session_state.employees.append(employee)

            st.success(
                f"{member_name} added successfully."
            )

            reset_form()

            st.rerun()


# ============================================================
# EMPLOYEE TABLE
# ============================================================

st.divider()

st.header("👥 Employee List")

employees = st.session_state.employees

if not employees:

    st.info(
        "No employees added yet. Add an employee using the form above."
    )

else:

    df = pd.DataFrame(employees)

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
    )

    st.caption(
        f"Total employees: {len(employees)}"
    )


# ============================================================
# SUMMARY
# ============================================================

if employees:

    st.divider()

    st.header("📊 Summary")

    total_gross = sum(
        employee["Gross Wages"]
        for employee in employees
    )

    total_epf = sum(
        employee["EPF Wages"]
        for employee in employees
    )

    total_ee = sum(
        employee["EE Share"]
        for employee in employees
    )

    total_eps = sum(
        employee["EPS Share"]
        for employee in employees
    )

    total_er = sum(
        employee["ER Share"]
        for employee in employees
    )

    total_refund = sum(
        employee["Refund of Advances"]
        for employee in employees
    )

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    c1.metric(
        "Employees",
        len(employees),
    )

    c2.metric(
        "Gross Wages",
        f"₹{total_gross:,}",
    )

    c3.metric(
        "EPF Wages",
        f"₹{total_epf:,}",
    )

    c4.metric(
        "EE Share",
        f"₹{total_ee:,}",
    )

    c5.metric(
        "EPS Share",
        f"₹{total_eps:,}",
    )

    c6.metric(
        "ER Share",
        f"₹{total_er:,}",
    )


# ============================================================
# EXPORT
# ============================================================

if employees:

    st.divider()

    st.header("⬇️ Export")

    ecr_text = generate_ecr_text(employees)

    excel_data = generate_excel(employees)

    col1, col2 = st.columns(2)

    with col1:

        st.download_button(
            label="📄 Download ECR TXT",
            data=ecr_text,
            file_name="EPFO_ECR.txt",
            mime="text/plain",
            use_container_width=True,
        )

    with col2:

        st.download_button(
            label="📊 Download Excel",
            data=excel_data,
            file_name="EPFO_ECR.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    with st.expander("👀 Preview ECR Text"):

        st.code(
            ecr_text,
            language="text",
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "⚠️ Before uploading the generated TXT file to EPFO, "
    "verify the current EPFO ECR file specification, field order, "
    "contribution rules and rounding requirements."
)
