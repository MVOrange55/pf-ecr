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

def clean_number(value):
    try:
        return int(float(value))
    except Exception:
        return 0


def calculate_employee(
    uan: str,
    name: str,
    gross: float,
    epf_wages: float,
    ncp_days: int,
    refund: float,
):
    """
    Calculate all ECR values for one employee.
    """

    eps_wages = min(epf_wages, EPS_WAGE_CEILING)
    edli_wages = min(epf_wages, EPS_WAGE_CEILING)

    # Employee EPF share = 12%
    ee_share = round(epf_wages * EPF_RATE)

    # Employer pension share = 8.33%
    eps_share = round(eps_wages * EPS_RATE)

    # Employer EPF share
    er_share = ee_share - eps_share

    # EDLI - informational
    edli_due = round(edli_wages * EDLI_RATE)

    return {
        "UAN": uan,
        "Member Name": name,
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
    errors = []
    warnings = []

    uan = str(uan).strip()
    name = str(name).strip()

    # UAN
    if not uan:
        errors.append("UAN is required.")

    elif not uan.isdigit():
        errors.append("UAN must contain digits only.")

    elif len(uan) != 12:
        errors.append(
            f"UAN must be exactly 12 digits. "
            f"Currently {len(uan)} digits."
        )

    # Name
    if not name:
        errors.append("Member name is required.")

    # Gross
    try:
        gross = float(gross)
    except Exception:
        errors.append("Gross wages must be a number.")
        gross = 0

    # EPF
    try:
        epf_wages = float(epf_wages)
    except Exception:
        errors.append("EPF wages must be a number.")
        epf_wages = 0

    if gross < 0:
        errors.append("Gross wages cannot be negative.")

    if epf_wages < 0:
        errors.append("EPF wages cannot be negative.")

    if epf_wages > gross:
        warnings.append(
            "EPF
