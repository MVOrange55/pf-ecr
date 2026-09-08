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
# HELPERS
# ============================================================

def clean_number(value: Any) -> int:
    """Convert a value to a non-negative integer."""

    try:
        return max(0, int(float(value)))
    except (ValueError, TypeError):
        return 0


def clean_uan(value: Any) -> str:
    """Clean UAN value safely, including Excel .0 values."""

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    text = str(value).strip()

    if text.endswith(".0"):
        text = text[:-2]

    return text


# ============================================================
# EMPLOYEE CALCULATION
# ============================================================

def calculate_employee(
    uan: str,
    name: str,
    gross: float,
    epf_wages: float,
    ncp_days: int,
    refund: float,
):
    """
    Calculate employee contribution values.
    """

    gross = max(0, float(gross))
    epf_wages = max(0, float(epf_wages))
    ncp_days = max(0, int(ncp_days))
    refund = max(0, float(refund))

    eps_wages = min(
        epf_wages,
        EPS_WAGE_CEILING,
    )

    edli_wages = min(
        epf_wages,
        EPS_WAGE_CEILING,
    )

    # Employee EPF share
    ee_share = int(
        round(epf_wages * EPF_RATE)
    )

    # Employer EPS share
    eps_share = int(
        round(eps_wages * EPS_RATE)
    )

    # Employer EPF portion
    er_share = max(
        0,
        ee_share - eps_share,
    )

    # EDLI
    edli_due = int(
        round(edli_wages * EDLI_RATE)
    )

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

    # UAN
    if not uan:
        errors.append(
            "UAN is required."
        )

    elif not uan.isdigit():
        errors.append(
            "UAN must contain digits only."
        )

    elif len(uan) != 12:
        errors.append(
            "UAN must be exactly 12 digits."
        )

    # Name
    if not name:
        errors.append(
            "Member name is required."
        )

    # Gross
    try:
        gross = float(gross)
    except (ValueError, TypeError):
        errors.append(
            "Gross wages must be a number."
        )
        gross = 0

    if gross < 0:
        errors.append(
            "Gross wages cannot be negative."
        )

    # EPF wages
    try:
        epf_wages = float(epf_wages)
    except (ValueError, TypeError):
        errors.append(
            "EPF wages must be a number."
        )
        epf_wages = 0

    if epf_wages < 0:
        errors.append(
            "EPF wages cannot be negative."
        )

    if epf_wages > gross:
        warnings.append(
            "EPF wages are greater than gross wages. "
            "Please verify the wages."
        )

    # NCP
    try:
        ncp_days = int(float(ncp_days))
    except (ValueError, TypeError):
        errors.append(
            "NCP days must be a number."
        )
        ncp_days = 0

    if ncp_days < 0:
        errors.append(
            "NCP days cannot be negative."
        )

    if ncp_days > 31:
        errors.append(
            "NCP days cannot be greater than 31."
        )

    # Refund
    try:
        refund = float(refund)
    except (ValueError, TypeError):
        errors.append(
            "Refund of Advances must be a number."
        )
        refund = 0

    if refund < 0:
        errors.append(
            "Refund of Advances cannot be negative."
        )

    return errors, warnings


# ============================================================
# ECR TXT GENERATOR
# ============================================================

def employee_to_ecr_line(
    employee: dict,
) -> str:
    """
    Generate ECR-style line.

    Verify exact current EPFO upload field order
    before production upload.
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

    return "|".join(
        str(value)
        for value in fields
    )


def generate_ecr_text(
    employees: list[dict],
) -> str:

    return "\n".join(
        employee_to_ecr_line(employee)
        for employee in employees
    )


# ============================================================
# EXCEL EXPORT
# ============================================================

def generate_excel(
    employees: list[dict],
) -> bytes:

    df = pd.DataFrame(
        employees
    )

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
# AUTOMATIC ACCOUNT SUMMARY
# ============================================================

def calculate_account_summary(
    employees: list[dict],
) -> dict:

    account_1 = sum(
        employee["EE Share"]
        for employee in employees
    )

    account_10 = sum(
        employee["EPS Share"]
        for employee in employees
    )

    employer_epf = sum(
        employee["ER Share"]
        for employee in employees
    )

    account_21 = sum(
        employee["EDLI Due"]
        for employee in employees
    )

    total_gross = sum(
        employee["Gross Wages"]
        for employee in employees
    )

    total_epf_wages = sum(
        employee["EPF Wages"]
        for employee in employees
    )

    total_eps_wages = sum(
        employee["EPS Wages"]
        for employee in employees
    )

    total_edli_wages = sum(
        employee["EDLI Wages"]
        for employee in employees
    )

    total_refund = sum(
        employee["Refund of Advances"]
        for employee in employees
    )

    total_contribution = (
        account_1
        + account_10
        + employer_epf
    )

    total_with_edli = (
        total_contribution
        + account_21
    )

    return {
        "employees": len(employees),
        "gross": total_gross,
        "epf_wages": total_epf_wages,
        "eps_wages": total_eps_wages,
        "edli_wages": total_edli_wages,
        "account_1": account_1,
        "account_10": account_10,
        "employer_epf": employer_epf,
        "account_21": account_21,
        "refund": total_refund,
        "total_contribution": total_contribution,
        "total_with_edli": total_with_edli,
    }


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
    reason_code = (
        str(reason_code)
        .strip()
        .upper()
    )

    if not uan:
        errors.append(
            "UAN is required."
        )

    elif not uan.isdigit():
        errors.append(
            "UAN must contain digits only."
        )

    elif len(uan) != 12:
        errors.append(
            "UAN must be exactly 12 digits."
        )

    if exit_date is None:
        errors.append(
            "Date of Exit is required."
        )

    if reason_code not in EXIT_REASONS.values():
        errors.append(
            "Invalid exit reason code."
        )

    return errors


# ============================================================
# BULK EXIT TXT
# ============================================================

def generate_bulk_exit_line(
    uan,
    exit_date,
    reason_code,
):
    """
    Bulk Exit format:

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
) -> str:

    return "\n".join(
        generate_bulk_exit_line(
            record["UAN"],
            record["Exit Date"],
            record["Reason Code"],
        )
        for record in exit_records
    )


# ============================================================
# FORM RESET
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

st.title(
    "📄 EPFO ECR & Bulk Exit Generator"
)

st.caption(
    "ECR + Account 1/10/21 Summary + "
    "Bulk Exit + Excel Import"
)

st.divider()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("⚙️ Settings")

    st.write("### Rates")

    st.write(
        f"EPF: **{EPF_RATE * 100:.2f}%**"
    )

    st.write(
        f"EPS: **{EPS_RATE * 100:.2f}%**"
    )

    st.write(
        f"EDLI: **{EDLI_RATE * 100:.2f}%**"
    )

    st.write(
        f"EPS Ceiling: **₹{EPS_WAGE_CEILING:,}**"
    )

    st.divider()

    st.write("### Records")

    st.write(
        f"Employees: **{len(st.session_state.employees)}**"
    )

    st.write(
        f"Exit Records: **{len(st.session_state.exit_records)}**"
    )

    st.divider()

    if st.button(
        "🗑️ Clear Employees",
        use_container_width=True,
    ):

        st.session_state.employees = []

        st.rerun()

    if st.button(
        "🗑️ Clear Exit Records",
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
# TAB 1
# ============================================================

with tab1:

    st.header("➕ Add Employee")

    with st.form(
        "employee_form"
    ):

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

            existing_uans = {
                employee["UAN"]
                for employee
                in st.session_state.employees
            }

            if clean_uan(uan) in existing_uans:

                st.error(
                    "This UAN already exists."
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
    # EMPLOYEE LIST
    # --------------------------------------------------------

    st.divider()

    st.subheader(
        "👥 Employee List"
    )

    employees = st.session_state.employees

    if not employees:

        st.info(
            "No employees added."
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
        # DELETE EMPLOYEE
        # ----------------------------------------------------

        options = [
            f"{i + 1}. "
            f"{employee['UAN']} - "
            f"{employee['Member Name']}"
            for i, employee
            in enumerate(employees)
        ]

        selected = st.selectbox(
            "Select employee to remove",
            options,
        )

        if st.button(
            "🗑️ Remove Employee",
            use_container_width=True,
        ):

            index = options.index(
                selected
            )

            st.session_state.employees.pop(
                index
            )

            st.rerun()

    # ========================================================
    # AUTOMATIC SUMMARY
    # ========================================================

    if employees:

        st.divider()

        st.header(
            "📊 Automatic PF Account Summary"
        )

        summary = calculate_account_summary(
            employees
        )

        # ----------------------------------------------------
        # TOP METRICS
        # ----------------------------------------------------

        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "👥 Employees",
            f"{summary['employees']:,}",
        )

        c2.metric(
            "💰 Gross Wages",
            f"₹{summary['gross']:,}",
        )

        c3.metric(
            "📌 EPF Wages",
            f"₹{summary['epf_wages']:,}",
        )

        c4.metric(
            "💵 Total Contribution",
            f"₹{summary['total_contribution']:,}",
        )

        st.divider()

        # ----------------------------------------------------
        # ACCOUNT CARDS
        # ----------------------------------------------------

        st.subheader(
            "🏦 Account-wise Amount"
        )

        a1, a10, a21 = st.columns(3)

        with a1:

            st.metric(
                "Account 1",
                f"₹{summary['account_1']:,}",
            )

            st.caption(
                "Employee EPF Share"
            )

        with a10:

            st.metric(
                "Account 10",
                f"₹{summary['account_10']:,}",
            )

            st.caption(
                "Employer EPS Share"
            )

        with a21:

            st.metric(
                "Account 21",
                f"₹{summary['account_21']:,}",
            )

            st.caption(
                "EDLI Contribution"
            )

        # ----------------------------------------------------
        # ACCOUNT TABLE
        # ----------------------------------------------------

        st.subheader(
            "📋 Account Summary"
        )

        account_df = pd.DataFrame(
            {
                "Account": [
                    "Account 1",
                    "Account 10",
                    "Account 21",
                ],
                "Description": [
                    "Employee EPF Share",
                    "Employer EPS Share",
                    "EDLI Contribution",
                ],
                "Amount": [
                    summary["account_1"],
                    summary["account_10"],
                    summary["account_21"],
                ],
            }
        )

        account_df["Amount"] = (
            account_df["Amount"]
            .apply(
                lambda x:
                f"₹{x:,}"
            )
        )

        st.dataframe(
            account_df,
            use_container_width=True,
            hide_index=True,
        )

        # ----------------------------------------------------
        # RECONCILIATION
        # ----------------------------------------------------

        st.subheader(
            "🧮 Reconciliation"
        )

        reconciliation_df = pd.DataFrame(
            {
                "Particular": [
                    "Account 1 - Employee EPF",
                    "Employer EPF Portion",
                    "Account 10 - EPS",
                    "Account 21 - EDLI",
                    "Total EPF + EPS",
                    "Total Including EDLI",
                    "Refund of Advances",
                ],
                "Amount": [
                    summary["account_1"],
                    summary["employer_epf"],
                    summary["account_10"],
                    summary["account_21"],
                    summary["total_contribution"],
                    summary["total_with_edli"],
                    summary["refund"],
                ],
            }
        )

        reconciliation_df["Amount"] = (
            reconciliation_df["Amount"]
            .apply(
                lambda x:
                f"₹{x:,}"
            )
        )

        st.dataframe(
            reconciliation_df,
            use_container_width=True,
            hide_index=True,
        )

        # ----------------------------------------------------
        # ECR EXPORT
        # ----------------------------------------------------

        st.divider()

        st.subheader(
            "⬇️ ECR Export"
        )

        ecr_text = generate_ecr_text(
            employees
        )

        excel_data = generate_excel(
            employees
        )

        c1, c2 = st.columns(2)

        with c1:

            st.download_button(
                "📄 Download ECR TXT",
                data=ecr_text,
                file_name="EPFO_ECR.txt",
                mime="text/plain",
                use_container_width=True,
            )

        with c2:

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
            "👀 ECR TXT Preview"
        ):

            st.code(
                ecr_text,
                language="text",
            )


# ============================================================
# TAB 2 - BULK EXIT
# ============================================================

with tab2:

    st.header(
        "🚪 Bulk Exit TXT Generator"
    )

    st.info(
        "Format: UAN#~#DD/MM/YYYY#~#Reason Code"
    )

    employees = st.session_state.employees
    exit_records = st.session_state.exit_records

    # --------------------------------------------------------
    # EXIT FORM
    # --------------------------------------------------------

    with st.form(
        "bulk_exit_form"
    ):

        col1, col2, col3 = st.columns(3)

        with col1:

            if employees:

                employee_uans = [
                    employee["UAN"]
                    for employee
                    in employees
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
                "Date of Exit"
            )

        with col3:

            reason_name = st.selectbox(
                "Reason for Exit",
                list(
                    EXIT_REASONS.keys()
                ),
            )

        reason_code = EXIT_REASONS[
            reason_name
        ]

        st.write(
            f"Reason Code: **{reason_code}**"
        )

        add_exit = st.form_submit_button(
            "➕ Add Exit Record",
            use_container_width=True,
            type="primary",
        )

    # --------------------------------------------------------
    # ADD EXIT
    # --------------------------------------------------------

    if add_exit:

        errors = validate_exit_record(
            exit_uan,
            exit_date,
            reason_code,
        )

        existing_exit_uans = {
            record["UAN"]
            for record in exit_records
        }

        if clean_uan(exit_uan) in existing_exit_uans:

            errors.append(
                "This UAN is already in "
                "the Bulk Exit list."
            )

        if errors:

            for error in errors:
                st.error(error)

        else:

            st.session_state.exit_records.append(
                {
                    "UAN": clean_uan(
                        exit_uan
                    ),
                    "Exit Date": exit_date,
                    "Reason": reason_name,
                    "Reason Code": reason_code,
                }
            )

            st.success(
                "Exit record added."
            )

            st.rerun()

    # --------------------------------------------------------
    # EXIT LIST
    # --------------------------------------------------------

    exit_records = st.session_state.exit_records

    if exit_records:

        st.divider()

        st.subheader(
            "📋 Exit Records"
        )

        exit_df = pd.DataFrame(
            exit_records
        )

        display_exit_df = exit_df.copy()

        display_exit_df[
            "Exit Date"
        ] = display_exit_df[
            "Exit Date"
        ].apply(
            lambda x:
            x.strftime("%d/%m/%Y")
        )

        st.dataframe(
            display_exit_df,
            use_container_width=True,
            hide_index=True,
        )

        # ----------------------------------------------------
        # REMOVE EXIT
        # ----------------------------------------------------

        remove_options = [
            f"{i + 1}. "
            f"{record['UAN']} - "
            f"{record['Exit Date'].strftime('%d/%m/%Y')}"
            for i, record
            in enumerate(exit_records)
        ]

        selected_exit = st.selectbox(
            "Select exit record to remove",
            remove_options,
        )

        if st.button(
            "🗑️ Remove Exit Record",
            use_container_width=True,
        ):

            index = remove_options.index(
                selected_exit
            )

            st.session_state.exit_records.pop(
                index
            )

            st.rerun()

        # ----------------------------------------------------
        # TXT
        # ----------------------------------------------------

        st.divider()

        st.subheader(
            "📄 Bulk Exit TXT"
        )

        bulk_exit_text = (
            generate_bulk_exit_text(
                exit_records
            )
        )

        st.text_area(
            "TXT Preview",
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

    else:

        st.info(
            "No exit records added."
        )


# ============================================================
# TAB 3 - BULK IMPORT
# ============================================================

with tab3:

    st.header(
        "📥 Bulk Employee Import"
    )

    st.write(
        "Excel/CSV file se multiple employees "
        "ek saath import karein."
    )

    st.info(
        "Required columns: "
        "UAN, Member Name, Gross Wages, "
        "EPF Wages, NCP Days, Refund of Advances"
    )

    uploaded_file = st.file_uploader(
        "Upload Excel / CSV",
        type=[
            "xlsx",
            "xls",
            "csv",
        ],
    )

    if uploaded_file:

        try:

            filename = (
                uploaded_file.name.lower()
            )

            if filename.endswith(
                ".csv"
            ):

                import_df = pd.read_csv(
                    uploaded_file,
                    dtype=str,
                )

            else:

                import_df = pd.read_excel(
                    uploaded_file,
                    dtype=str,
                )

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
                if column
                not in import_df.columns
            ]

            if missing_columns:

                st.error(
                    "Missing columns: "
                    + ", ".join(
                        missing_columns
                    )
                )

                st.code(
                    ", ".join(
                        required_columns
                    )
                )

            else:

                if st.button(
                    "📥 Import Employees",
                    use_container_width=True,
                    type="primary",
                ):

                    imported = 0
                    skipped = 0

                    existing_uans = {
                        employee["UAN"]
                        for employee
                        in st.session_state.employees
                    }

                    import_errors = []

                    for row_number, row in import_df.iterrows():

                        excel_row = (
                            row_number + 2
                        )

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

                            skipped += 1

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

                        imported += 1

                    st.success(
                        f"{imported} employees imported."
                    )

                    if skipped:

                        st.warning(
                            f"{skipped} rows skipped."
                        )

                        with st.expander(
                            "View Import Errors"
                        ):

                            for error in import_errors:
                                st.error(error)

                    if imported:

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
    "⚠️ Verify the generated ECR/Bulk Exit file against "
    "the current EPFO portal specification before upload."
)
