from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="EPFO ECR & Bulk Exit Generator",
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

# EPFO Bulk Exit reason codes
EXIT_REASONS = {
    "Retirement": "R",
    "Death in Service": "D",
    "Superannuation": "S",
    "Permanent Disablement": "P",
    "Cessation (Short Service)": "C",
    "Death Away From Service": "A",
}


# ============================================================
# SESSION STATE
# ============================================================

if "employees" not in st.session_state:
    st.session_state.employees = []

if "exit_records" not in st.session_state:
    st.session_state.exit_records = []


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_number(value: Any) -> int:
    """Convert value to a non-negative integer."""

    try:
        return max(0, int(float(value)))
    except (ValueError, TypeError):
        return 0


def clean_uan(value: Any) -> str:
    """
    Clean UAN without accidentally converting it to scientific notation.
    """

    if pd.isna(value):
        return ""

    text = str(value).strip()

    # Handle Excel values like 100248330106.0
    if text.endswith(".0"):
        text = text[:-2]

    return text


# ============================================================
# ECR CALCULATION
# ============================================================

def calculate_employee(
    uan: str,
    name: str,
    gross: float,
    epf_wages: float,
    ncp_days: int,
    refund: float,
):
    """Calculate ECR contribution values."""

    gross = max(0, float(gross))
    epf_wages = max(0, float(epf_wages))
    ncp_days = max(0, int(ncp_days))
    refund = max(0, float(refund))

    eps_wages = min(epf_wages, EPS_WAGE_CEILING)
    edli_wages = min(epf_wages, EPS_WAGE_CEILING)

    # Employee EPF share
    ee_share = int(round(epf_wages * EPF_RATE))

    # Employer EPS share
    eps_share = int(round(eps_wages * EPS_RATE))

    # Employer EPF share
    er_share = max(0, ee_share - eps_share)

    # EDLI - informational
    edli_due = int(round(edli_wages * EDLI_RATE))

    return {
        "UAN": clean_uan(uan),
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


# ============================================================
# EMPLOYEE VALIDATION
# ============================================================

def validate_employee(
    uan,
    name,
    gross,
    epf_wages,
    ncp_days,
    refund,
):
    errors = []
    warnings = []

    uan = clean_uan(uan)
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
    # NAME
    # --------------------------------------------------------

    if not name:
        errors.append("Member name is required.")

    # --------------------------------------------------------
    # GROSS
    # --------------------------------------------------------

    try:
        gross = float(gross)
    except (ValueError, TypeError):
        errors.append("Gross wages must be a number.")
        gross = 0

    if gross < 0:
        errors.append("Gross wages cannot be negative.")

    # --------------------------------------------------------
    # EPF WAGES
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
            "EPF wages are greater than gross wages. "
            "Please verify the wages."
        )

    # --------------------------------------------------------
    # NCP DAYS
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
    # REFUND
    # --------------------------------------------------------

    try:
        refund = float(refund)
    except (ValueError, TypeError):
        errors.append("Refund of Advances must be a number.")
        refund = 0

    if refund < 0:
        errors.append(
            "Refund of Advances cannot be negative."
        )

    return errors, warnings


# ============================================================
# ECR TXT
# ============================================================

def employee_to_ecr_line(employee: dict) -> str:
    """
    Generate ECR-style line.

    IMPORTANT:
    Verify exact field sequence against the current EPFO
    upload specification before portal submission.
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

    return "|".join(str(value) for value in fields)


def generate_ecr_text(employees: list[dict]) -> str:

    return "\n".join(
        employee_to_ecr_line(employee)
        for employee in employees
    )


# ============================================================
# EXCEL EXPORT
# ============================================================

def generate_excel(employees: list[dict]) -> bytes:

    df = pd.DataFrame(employees)

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="ECR Data",
        )

    return output.getvalue()


# ============================================================
# BULK EXIT VALIDATION
# ============================================================

def validate_exit_record(
    uan,
    exit_date,
    reason_code,
):
    errors = []

    uan = clean_uan(uan)
    reason_code = str(reason_code).strip().upper()

    # UAN
    if not uan:
        errors.append("UAN is required.")

    elif not uan.isdigit():
        errors.append(
            "UAN must contain digits only."
        )

    elif len(uan) != 12:
        errors.append(
            "UAN must be exactly 12 digits."
        )

    # Exit date
    if exit_date is None:
        errors.append(
            "Date of Exit is required."
        )

    # Reason
    if reason_code not in EXIT_REASONS.values():
        errors.append(
            "Invalid exit reason code."
        )

    return errors


# ============================================================
# BULK EXIT LINE
# ============================================================

def generate_bulk_exit_line(
    uan,
    exit_date,
    reason_code,
):
    """
    EPFO Bulk Exit structure:

    UAN#~#DD/MM/YYYY#~#Reason Code
    """

    return (
        f"{clean_uan(uan)}"
        f"#~#"
        f"{exit_date.strftime('%d/%m/%Y')}"
        f"#~#"
        f"{str(reason_code).strip().upper()}"
    )


def generate_bulk_exit_text(
    exit_records: list[dict],
):

    return "\n".join(
        generate_bulk_exit_line(
            record["UAN"],
            record["Exit Date"],
            record["Reason Code"],
        )
        for record in exit_records
    )


# ============================================================
# RESET FUNCTIONS
# ============================================================

def reset_employee_form():

    st.session_state.uan = ""
    st.session_state.member_name = ""
    st.session_state.gross_wages = 0
    st.session_state.epf_wages = 0
    st.session_state.ncp_days = 0
    st.session_state.refund = 0


# ============================================================
# HEADER
# ============================================================

st.title("📄 EPFO ECR & Bulk Exit Generator")

st.caption(
    "ECR calculation + employee management + "
    "Bulk Exit TXT generation"
)

st.divider()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Settings")

    st.write("### Contribution Rates")

    st.write(
        f"EPF Rate: **{EPF_RATE * 100:.2f}%**"
    )

    st.write(
        f"EPS Rate: **{EPS_RATE * 100:.2f}%**"
    )

    st.write(
        f"EDLI Rate: **{EDLI_RATE * 100:.2f}%**"
    )

    st.write(
        f"EPS Wage Ceiling: **₹{EPS_WAGE_CEILING:,}**"
    )

    st.divider()

    st.write("### Current Records")

    st.write(
        f"Employees: **{len(st.session_state.employees)}**"
    )

    st.write(
        f"Exit Records: **{len(st.session_state.exit_records)}**"
    )

    st.divider()

    if st.button(
        "🗑️ Clear All Employees",
        use_container_width=True,
    ):

        st.session_state.employees = []

        st.rerun()

    if st.button(
        "🗑️ Clear All Exit Records",
        use_container_width=True,
    ):

        st.session_state.exit_records = []

        st.rerun()


# ============================================================
# TABS
# ============================================================

tab1, tab2, tab3 = st.tabs(
    [
        "👥 ECR Employees",
        "🚪 Bulk Exit",
        "📥 Bulk Import",
    ]
)


# ============================================================
# TAB 1 - ECR EMPLOYEES
# ============================================================

with tab1:

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

    # --------------------------------------------------------
    # ADD EMPLOYEE
    # --------------------------------------------------------

    if submitted:

        errors, warnings = validate_employee(
            uan,
            member_name,
            gross_wages,
            epf_wages,
            ncp_days,
            refund,
        )

        if errors:

            for error in errors:
                st.error(error)

        else:

            for warning in warnings:
                st.warning(warning)

            existing_uans = [
                employee["UAN"]
                for employee in st.session_state.employees
            ]

            if clean_uan(uan) in existing_uans:

                st.error(
                    f"UAN {uan} already exists."
                )

            else:

                employee = calculate_employee(
                    uan,
                    member_name,
                    gross_wages,
                    epf_wages,
                    ncp_days,
                    refund,
                )

                st.session_state.employees.append(
                    employee
                )

                st.success(
                    f"{member_name} added successfully."
                )

                reset_employee_form()

                st.rerun()

    # --------------------------------------------------------
    # EMPLOYEE TABLE
    # --------------------------------------------------------

    st.divider()

    st.subheader("👥 Employee List")

    employees = st.session_state.employees

    if not employees:

        st.info(
            "No employees added yet."
        )

    else:

        employee_df = pd.DataFrame(
            employees
        )

        st.dataframe(
            employee_df,
            use_container_width=True,
            hide_index=True,
        )

        st.caption(
            f"Total employees: {len(employees)}"
        )

        # ----------------------------------------------------
        # REMOVE EMPLOYEE
        # ----------------------------------------------------

        st.subheader("🗑️ Remove Employee")

        remove_employee_options = [
            f"{index + 1}. "
            f"{employee['UAN']} - "
            f"{employee['Member Name']}"
            for index, employee
            in enumerate(employees)
        ]

        selected_employee = st.selectbox(
            "Select employee",
            remove_employee_options,
        )

        if st.button(
            "🗑️ Remove Selected Employee",
            use_container_width=True,
        ):

            index = (
                remove_employee_options.index(
                    selected_employee
                )
            )

            st.session_state.employees.pop(
                index
            )

            st.success(
                "Employee removed successfully."
            )

            st.rerun()

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    if employees:

        st.divider()

        st.subheader("📊 ECR Summary")

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

        c1, c2, c3 = st.columns(3)
        c4, c5, c6 = st.columns(3)

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

        # ----------------------------------------------------
        # ECR EXPORT
        # ----------------------------------------------------

        st.divider()

        st.subheader("⬇️ ECR Export")

        ecr_text = generate_ecr_text(
            employees
        )

        excel_data = generate_excel(
            employees
        )

        col1, col2 = st.columns(2)

        with col1:

            st.download_button(
                "📄 Download ECR TXT",
                data=ecr_text,
                file_name="EPFO_ECR.txt",
                mime="text/plain",
                use_container_width=True,
            )

        with col2:

            st.download_button(
                "📊 Download Excel",
                data=excel_data,
                file_name="EPFO_ECR.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

        with st.expander(
            "👀 Preview ECR TXT"
        ):

            st.code(
                ecr_text,
                language="text",
            )


# ============================================================
# TAB 2 - BULK EXIT
# ============================================================

with tab2:

    st.header("🚪 Bulk Exit TXT Generator")

    st.info(
        "Format: UAN#~#DD/MM/YYYY#~#Reason Code"
    )

    exit_records = st.session_state.exit_records

    # --------------------------------------------------------
    # ADD EXIT RECORD
    # --------------------------------------------------------

    with st.form("bulk_exit_form"):

        col1, col2, col3 = st.columns(3)

        with col1:

            if employees:

                employee_uans = [
                    employee["UAN"]
                    for employee in employees
                ]

                exit_uan = st.selectbox(
                    "Select UAN",
                    employee_uans,
                )

            else:

                exit_uan = st.text_input(
                    "UAN",
                    max_chars=12,
                    placeholder="12 digit UAN",
                )

        with col2:

            exit_date = st.date_input(
                "Date of Exit",
            )

        with col3:

            reason_name = st.selectbox(
                "Reason for Exit",
                list(EXIT_REASONS.keys()),
            )

        reason_code = EXIT_REASONS[
            reason_name
        ]

        st.caption(
            f"Selected reason code: **{reason_code}**"
        )

        add_exit = st.form_submit_button(
            "➕ Add Exit Record",
            use_container_width=True,
            type="primary",
        )

    # --------------------------------------------------------
    # SAVE EXIT
    # --------------------------------------------------------

    if add_exit:

        errors = validate_exit_record(
            exit_uan,
            exit_date,
            reason_code,
        )

        existing_exit_uans = [
            record["UAN"]
            for record in exit_records
        ]

        if clean_uan(exit_uan) in existing_exit_uans:

            errors.append(
                "This UAN already exists in the "
                "Bulk Exit list."
            )

        if errors:

            for error in errors:
                st.error(error)

        else:

            st.session_state.exit_records.append(
                {
                    "UAN": clean_uan(exit_uan),
                    "Exit Date": exit_date,
                    "Reason": reason_name,
                    "Reason Code": reason_code,
                }
            )

            st.success(
                f"Exit record added for UAN {exit_uan}."
            )

            st.rerun()

    # --------------------------------------------------------
    # EXIT TABLE
    # --------------------------------------------------------

    exit_records = st.session_state.exit_records

    if exit_records:

        st.divider()

        st.subheader("📋 Exit Records")

        exit_df = pd.DataFrame(
            exit_records
        )

        display_df = exit_df.copy()

        display_df["Exit Date"] = (
            display_df["Exit Date"].apply(
                lambda value:
                value.strftime("%d/%m/%Y")
            )
        )

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
        )

        # ----------------------------------------------------
        # REMOVE EXIT
        # ----------------------------------------------------

        st.subheader(
            "🗑️ Remove Exit Record"
        )

        remove_exit_options = [
            f"{index + 1}. "
            f"{record['UAN']} - "
            f"{record['Exit Date'].strftime('%d/%m/%Y')}"
            for index, record
            in enumerate(exit_records)
        ]

        selected_exit = st.selectbox(
            "Select exit record",
            remove_exit_options,
        )

        if st.button(
            "🗑️ Remove Selected Exit",
            use_container_width=True,
        ):

            index = (
                remove_exit_options.index(
                    selected_exit
                )
            )

            st.session_state.exit_records.pop(
                index
            )

            st.success(
                "Exit record removed."
            )

            st.rerun()

        # ----------------------------------------------------
        # BULK EXIT TXT
        # ----------------------------------------------------

        st.divider()

        st.subheader(
            "📄 Bulk Exit TXT Preview"
        )

        bulk_exit_text = (
            generate_bulk_exit_text(
                exit_records
            )
        )

        st.text_area(
            "Generated TXT",
            value=bulk_exit_text,
            height=300,
        )

        st.download_button(
            "⬇️ Download Bulk Exit TXT",
            data=bulk_exit_text,
            file_name="EPFO_Bulk_Exit.txt",
            mime="text/plain",
            use_container_width=True,
            type="primary",
        )

        st.caption(
            "Each line: UAN#~#DD/MM/YYYY#~#Reason Code"
        )

    else:

        st.info(
            "No exit records added yet."
        )


# ============================================================
# TAB 3 - BULK IMPORT
# ============================================================

with tab3:

    st.header("📥 Bulk Employee Import")

    st.write(
        "Excel/CSV file se multiple employees "
        "ek saath import karein."
    )

    st.info(
        "Recommended columns: "
        "UAN, Member Name, Gross Wages, "
        "EPF Wages, NCP Days, Refund of Advances"
    )

    uploaded_file = st.file_uploader(
        "Upload Excel or CSV",
        type=["xlsx", "xls", "csv"],
    )

    if uploaded_file:

        try:

            file_name = uploaded_file.name.lower()

            if file_name.endswith(".csv"):

                import_df = pd.read_csv(
                    uploaded_file,
                    dtype=str,
                )

            else:

                import_df = pd.read_excel(
                    uploaded_file,
                    dtype=str,
                )

            # Normalize column names
            import_df.columns = [
                str(column).strip()
                for column in import_df.columns
            ]

            st.subheader(
                "📋 Uploaded Data"
            )

            st.dataframe(
                import_df,
                use_container_width=True,
                hide_index=True,
            )

            required_columns = [
                "UAN",
                "Member Name",
                "Gross Wages",
                "EPF Wages",
                "NCP Days",
                "Refund of Advances",
            ]

            missing_columns = [
                column
                for column in required_columns
                if column not in import_df.columns
            ]

            if missing_columns:

                st.error(
                    "Missing columns: "
                    + ", ".join(
                        missing_columns
                    )
                )

                st.write(
                    "Required columns:"
                )

                st.code(
                    ", ".join(
                        required_columns
                    )
                )

            else:

                if st.button(
                    "📥 Import Employees",
                    type="primary",
                    use_container_width=True,
                ):

                    imported_count = 0
                    skipped_count = 0

                    existing_uans = {
                        employee["UAN"]
                        for employee
                        in st.session_state.employees
                    }

                    import_errors = []

                    for row_number, row in import_df.iterrows():

                        excel_row = row_number + 2

                        row_uan = clean_uan(
                            row["UAN"]
                        )

                        row_name = str(
                            row["Member Name"]
                        ).strip()

                        row_gross = row[
                            "Gross Wages"
                        ]

                        row_epf = row[
                            "EPF Wages"
                        ]

                        row_ncp = row[
                            "NCP Days"
                        ]

                        row_refund = row[
                            "Refund of Advances"
                        ]

                        errors, warnings = (
                            validate_employee(
                                row_uan,
                                row_name,
                                row_gross,
                                row_epf,
                                row_ncp,
                                row_refund,
                            )
                        )

                        if row_uan in existing_uans:

                            errors.append(
                                "Duplicate UAN."
                            )

                        if errors:

                            skipped_count += 1

                            import_errors.append(
                                f"Row {excel_row}: "
                                + " ".join(errors)
                            )

                            continue

                        employee = (
                            calculate_employee(
                                row_uan,
                                row_name,
                                row_gross,
                                row_epf,
                                row_ncp,
                                row_refund,
                            )
                        )

                        st.session_state.employees.append(
                            employee
                        )

                        existing_uans.add(
                            row_uan
                        )

                        imported_count += 1

                    st.success(
                        f"{imported_count} employees imported."
                    )

                    if skipped_count:

                        st.warning(
                            f"{skipped_count} rows skipped."
                        )

                        with st.expander(
                            "View Import Errors"
                        ):

                            for error in import_errors:
                                st.error(error)

                    if imported_count:

                        st.rerun()

        except Exception as error:

            st.error(
                f"File read error: {error}"
            )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "⚠️ Generated files should be checked against the "
    "current EPFO portal specification before upload."
)

st.caption(
    "EPFO ECR and Bulk Exit formats/rules can change. "
    "Always verify the latest official EPFO requirements."
)
