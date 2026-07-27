"""
Census Validator & Reformatter - Streamlit App
==============================================
Install:  pip install streamlit pandas zipcodes openpyxl rapidfuzz
Run:      streamlit run census_validator_app_v4_fixed.py

Purpose
-------
Validates and reformats CSA census files into the required 30-column import
layout. Handles standard vertical census files, horizontal dependent census
files, files with extra rows above the header, and common employer-export
headers with line breaks or parenthetical instructions.

"""

from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass
from typing import Any, Optional

import pandas as pd

try:  # Streamlit is required for the UI, but a tiny fallback helps local import tests.
    import streamlit as st
except ImportError:  # pragma: no cover
    class _StreamlitFallback:
        @staticmethod
        def cache_data(show_spinner: bool = False):
            def decorator(func):
                return func
            return decorator

        def __getattr__(self, name: str):
            raise RuntimeError("Streamlit is not installed. Install streamlit to run the app UI.")

    st = _StreamlitFallback()  # type: ignore

try:
    import zipcodes as _zc
    ZIPCODES_AVAILABLE = True
except ImportError:  # pragma: no cover
    ZIPCODES_AVAILABLE = False

try:
    from rapidfuzz import fuzz, process as rf_process
    FUZZY_AVAILABLE = True
except ImportError:  # pragma: no cover
    FUZZY_AVAILABLE = False


# =============================================================================
# 1. CANONICAL SCHEMA
# =============================================================================

SUPPORTED_ZIP_STATES = {
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA",
    "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT",
    "VA", "WA", "WV", "WI", "WY", "DC",
}

KNOWN_INVALID_OR_UNSUPPORTED_ZIPS = {
    "00820",
    "03333",
    "33210",
    "57507",
    "85403",
}

CLEARED_FIELDS = {
    "Employee ID", "Gender", "Email",
    "Address Line 1", "Address Line 2", "City", "State",
}

FIELDS: list[dict[str, Any]] = [
    {"col": "A",  "name": "First Name", "required": True, "type": "alpha"},
    {"col": "B",  "name": "Last Name", "required": True, "type": "alpha"},
    {"col": "C",  "name": "Employee ID", "required": False, "type": "alpha"},
    {"col": "D",  "name": "Relationship", "required": True, "type": "relationship"},
    {"col": "E",  "name": "DOB", "required": True, "type": "date"},
    {"col": "F",  "name": "Gender", "required": False, "type": "alpha"},
    {"col": "G",  "name": "Email", "required": False, "type": "alpha"},
    {"col": "H",  "name": "Address Line 1", "required": False, "type": "alpha"},
    {"col": "I",  "name": "Address Line 2", "required": False, "type": "alpha"},
    {"col": "J",  "name": "City", "required": False, "type": "alpha"},
    {"col": "K",  "name": "State", "required": False, "type": "alpha"},
    {"col": "L",  "name": "Zip Code", "required": True, "type": "zipcode"},
    {"col": "M",  "name": "County", "required": False, "type": "alpha"},
    {"col": "N",  "name": "Primary Worksite Zip Code", "required": True, "type": "zipcode"},
    {"col": "O",  "name": "Primary Worksite County", "required": False, "type": "alpha"},
    {"col": "P",  "name": "ICHRA Class", "required": False, "type": "alpha"},
    {"col": "Q",  "name": "Health Election", "required": True, "type": "election"},
    {"col": "R",  "name": "Current Health Plan Vendor", "required": False, "type": "alpha"},
    {"col": "S",  "name": "Current Health Plan", "required": False, "type": "alpha"},
    {"col": "T",  "name": "Current Health Plan Tier", "required": False, "type": "tier"},
    {"col": "U",  "name": "Current Health Plan OOP (single)", "required": False, "type": "numeric"},
    {"col": "V",  "name": "Current Health Plan OOP (family)", "required": False, "type": "numeric"},
    {"col": "W",  "name": "Current Health Plan Deductible (single)", "required": False, "type": "numeric"},
    {"col": "X",  "name": "Current Health Plan Deductible (family)", "required": False, "type": "numeric"},
    {"col": "Y",  "name": "Current Health Plan ER Cost", "required": False, "type": "numeric"},
    {"col": "Z",  "name": "Current Health Plan EE Cost", "required": False, "type": "numeric"},
    {"col": "AA", "name": "Annual Salary", "required": False, "type": "numeric"},
    {"col": "AB", "name": "Hourly Rate", "required": False, "type": "numeric"},
    {"col": "AC", "name": "Hours Per Week", "required": False, "type": "numeric"},
    {"col": "AD", "name": "Notes", "required": False, "type": "alpha"},
]

OUTPUT_COLUMNS = [f["name"] for f in FIELDS]
FIELD_BY_NAME = {f["name"]: f for f in FIELDS}
FIELD_IDX_BY_NAME = {f["name"]: i for i, f in enumerate(FIELDS)}
REQUIRED_FIELDS = {f["name"] for f in FIELDS if f["required"]}


# =============================================================================
# 2. NORMALIZATION HELPERS
# =============================================================================

def _norm(value: object) -> str:
    value = str(value or "").strip().lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _compact(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def _is_blank(value: object) -> bool:
    return str(value or "").strip() == ""


ANCILLARY_TERMS = {
    "dental", "vision", "life", "disability", "std", "ltd", "accident", "critical", "illness",
    "hospital", "indemnity", "cancer", "legal", "pet", "commuter", "parking", "transit",
    "supplemental", "voluntary", "add", "ad d",
}

HEALTH_TERMS = {
    "health", "medical", "med", "current health", "plan name", "product name", "carrier",
    "coverage tier", "coverage level", "benefit tier", "election tier", "employee cost",
    "employer cost", "ee cost", "er cost", "waiver", "waive",
}

HEADER_ALIASES: dict[str, str] = {
    # First / Last name    
    "first name": "First Name",
    "firstname": "First Name",
    "first": "First Name",
    "given name": "First Name",
    "employee first name": "First Name",
    "ee first name": "First Name",
    "member first name": "First Name",
    "subscriber first name": "First Name",
    
    "last name": "Last Name",
    "lastname": "Last Name",
    "last": "Last Name",
    "surname": "Last Name",
    "family name": "Last Name",
    "employee last name": "Last Name",
    "ee last name": "Last Name",
    "member last name": "Last Name",
    "subscriber last name": "Last Name",
    # Employee ID
    "employee id": "Employee ID",
    "emp id": "Employee ID",
    "empid": "Employee ID",
    "employee number": "Employee ID",
    "emp number": "Employee ID",
    "employee #": "Employee ID",
    "worker id": "Employee ID",
    "person number": "Employee ID",
    "ee #": "Employee ID",
    "ee not ssn": "Employee ID",
    "ee # not ssn": "Employee ID",
    # Relationship
    "relationship": "Relationship",
    "relation": "Relationship",
    "relationship to employee": "Relationship",
    "member type": "Relationship",
    "covered person type": "Relationship",
    # DOB
    "dob": "DOB",
    "date of birth": "DOB",
    "birth date": "DOB",
    "birthdate": "DOB",
    "employee dob": "DOB",
    "employee date of birth": "DOB",
    # gender
    "gender": "Gender",
    "sex": "Gender",
    "gender m f": "Gender",
    "gender male female": "Gender",
    # Email
    "email": "Email",
    "email address": "Email",
    "e mail": "Email",
    "work email": "Email",
    "personal email": "Email",
    # Address
    "address line 1": "Address Line 1",
    "address1": "Address Line 1",
    "street address": "Address Line 1",
    "primary address line 1": "Address Line 1",
    "home address line 1": "Address Line 1",
    "address": "Address Line 1",
    "address line 2": "Address Line 2",
    "address2": "Address Line 2",
    "apt": "Address Line 2",
    "suite": "Address Line 2",
    "primary address line 2": "Address Line 2",
    # City / State / Zip / County
    "city": "City",
    "town": "City",
    "primary address city": "City",
    "home city": "City",
    "primary address city town": "City",
    "state": "State",
    "st": "State",
    "state code": "State",
    "primary address state": "State",
    "primary address state territory": "State",
    "home state": "State",

    "zip code": "Zip Code",
    "zip": "Zip Code",
    "zipcode": "Zip Code",
    "postal code": "Zip Code",
    "home zip": "Zip Code",
    "home zip code": "Zip Code",
    "primary address zip code": "Zip Code",
    "primary address postal code": "Zip Code",

    "county": "County",
    "home county": "County",
    "primary address county": "County",
    # Worksite
    "primary worksite zip code": "Primary Worksite Zip Code",
    "worksite zip": "Primary Worksite Zip Code",
    "worksite zip code": "Primary Worksite Zip Code",
    "work zip": "Primary Worksite Zip Code",
    "work zip code": "Primary Worksite Zip Code",
    "site zip": "Primary Worksite Zip Code",
    "work location zip": "Primary Worksite Zip Code",
    "work location zip code": "Primary Worksite Zip Code",

    "primary worksite county": "Primary Worksite County",
    "worksite county": "Primary Worksite County",
    "work county": "Primary Worksite County",
    "work location county": "Primary Worksite County",
    # Class
    "ichra class": "ICHRA Class",
    "ichra": "ICHRA Class",
    "class": "ICHRA Class",
    "benefit class": "ICHRA Class",
    "worker category": "ICHRA Class",
    "employee class": "ICHRA Class",
    "employee class name or #": "ICHRA Class",
    "employee class name or": "ICHRA Class",
    # Health Election
    "health election": "Health Election",
    "election": "Health Election",
    "coverage election": "Health Election",
    "medical": "Health Election",
    "medical coverage": "Health Election",
    "medical coverage code": "Health Election",
    "medical election": "Health Election",
    "benefit election": "Health Election",
    "medical coverage election": "Health Election",
    # Health Plan Fields
    "current health vendor": "Current Health Plan Vendor",
    "current health plan vendor": "Current Health Plan Vendor",
    "health vendor": "Current Health Plan Vendor",
    "medical vendor": "Current Health Plan Vendor",
    "carrier": "Current Health Plan Vendor",
    "insurance carrier": "Current Health Plan Vendor",
    "health carrier": "Current Health Plan Vendor",
    "medical carrier": "Current Health Plan Vendor",

    "current health plan": "Current Health Plan",
    "health plan": "Current Health Plan",
    "plan name": "Current Health Plan",
    "medical plan": "Current Health Plan",
    "medical plan name": "Current Health Plan",
    "health plan name": "Current Health Plan",
    "current plan": "Current Health Plan",
    "product name": "Current Health Plan",
    "medical product name": "Current Health Plan",

    "current health plan tier": "Current Health Plan Tier",
    "health plan tier": "Current Health Plan Tier",
    "medical plan tier": "Current Health Plan Tier",
    "coverage tier": "Current Health Plan Tier",
    "coverage level": "Current Health Plan Tier",
    "plan tier": "Current Health Plan Tier",
    "tier": "Current Health Plan Tier",
    "medical coverage tier": "Current Health Plan Tier",
    "medical coverage level": "Current Health Plan Tier",

    "current health plan oop single": "Current Health Plan OOP (single)",
    "health plan oop single": "Current Health Plan OOP (single)",
    "oop single": "Current Health Plan OOP (single)",
    "out of pocket single": "Current Health Plan OOP (single)",
    "current health plan oop family": "Current Health Plan OOP (family)",
    "health plan oop family": "Current Health Plan OOP (family)",
    "oop family": "Current Health Plan OOP (family)",
    "out of pocket family": "Current Health Plan OOP (family)",
    "current health plan deductible single": "Current Health Plan Deductible (single)",
    "health plan deductible single": "Current Health Plan Deductible (single)",
    "deductible single": "Current Health Plan Deductible (single)",
    "current health plan deductible family": "Current Health Plan Deductible (family)",
    "health plan deductible family": "Current Health Plan Deductible (family)",
    "deductible family": "Current Health Plan Deductible (family)",

    "current health plan er cost": "Current Health Plan ER Cost",
    "health plan er cost": "Current Health Plan ER Cost",
    "medical er cost": "Current Health Plan ER Cost",
    "er cost": "Current Health Plan ER Cost",
    "employer cost": "Current Health Plan ER Cost",
    "employer contribution": "Current Health Plan ER Cost",
    "pending employer cost": "Current Health Plan ER Cost",
    "current health plan ee cost": "Current Health Plan EE Cost",
    "health plan ee cost": "Current Health Plan EE Cost",
    "medical ee cost": "Current Health Plan EE Cost",
    "ee cost": "Current Health Plan EE Cost",
    "employee cost": "Current Health Plan EE Cost",
    "employee contribution": "Current Health Plan EE Cost",
    "pending employee cost": "Current Health Plan EE Cost",
    # Compensation
    "annual salary": "Annual Salary",
    "gross annual earnings": "Annual Salary",
    "salary": "Annual Salary",
    "yearly salary": "Annual Salary",
    "base salary": "Annual Salary",
    "hourly rate": "Hourly Rate",
    "hourly": "Hourly Rate",
    "rate": "Hourly Rate",
    "hourly wage": "Hourly Rate",
    "hours per week": "Hours Per Week",
    "hours week": "Hours Per Week",
    "weekly hours": "Hours Per Week",
    "weekly hours worked": "Hours Per Week",
    "hrs per week": "Hours Per Week",
    # Notes
    "notes": "Notes",
    "note": "Notes",
    "comments": "Notes",
    "comment": "Notes",
    "remarks": "Notes",
    "medical waiver reason": "Notes",
}

for col, field_def in zip(
    ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD"],
    FIELDS,
):
    HEADER_ALIASES[col.lower()] = field_def["name"]


def _is_ancillary_header(header: object) -> bool:
    h = _norm(header)
    words = set(h.split())
    return any(term in words or term in h for term in ANCILLARY_TERMS)


def _looks_health_header(header: object) -> bool:
    h = _norm(header)
    if _is_ancillary_header(h):
        return False
    return any(term in h for term in HEALTH_TERMS) or "medical" in h or "med " in f"{h} "


def _is_verbose_vertical_relationship_header(header: object) -> bool:
    h = _norm(header)
    if h in {"relationship", "relation", "relationship to employee"}:
        return True
    if not h.startswith("relationship") and " relationship " not in f" {h} ":
        return False
    has_employee_token = bool(re.search(r"\b(employee|ee|emp|subscriber|self)\b", h))
    has_spouse_token = bool(re.search(r"\b(spouse|sp|sps|partner)\b", h))
    has_child_token = bool(re.search(r"\b(child|children|ch|dep|dependent)\b", h))
    starts_as_relationship_definition = h.startswith("relationship employee") or h.startswith("relationship ee")
    return starts_as_relationship_definition or (has_employee_token and has_spouse_token and has_child_token)


def _is_actual_dependent_header(header: object) -> bool:
    h = _norm(header)
    if not h:
        return False
    if _is_verbose_vertical_relationship_header(h):
        return False

    # Some employer census templates use abbreviated child headers like
    # "CH #1 DOB" / "CH#1 GENDER" rather than spelling out "Child".
    # Treat those as dependent-specific headers only when they include a
    # dependent attribute, so unrelated words containing "ch" do not trigger
    # horizontal conversion.
    has_dependent_word = bool(re.search(r"\b(dependent|dependant|dep)\b", h))
    has_spouse_child_word = bool(re.search(r"\b(spouse|child|children|domestic partner)\b", h))
    has_child_abbreviation = bool(re.search(r"\bch\b", h))
    has_dependent_attribute = any(token in h for token in ["dob", "birth", "first", "last", "name", "relationship", "relation", "type", "gender"])

    if has_dependent_word and has_dependent_attribute:
        return True
    if (has_spouse_child_word or has_child_abbreviation) and any(token in h for token in ["dob", "birth", "first", "last", "name", "gender"]):
        return True
    if (has_spouse_child_word or has_child_abbreviation) and (h.endswith("relationship") or h.endswith("relation") or h.endswith("type")):
        return True
    return False

def canonical_from_header(header: object) -> Optional[str]:
    h = _norm(header)
    if not h:
        return None
    if h in HEADER_ALIASES:
        return HEADER_ALIASES[h]
    if _is_verbose_vertical_relationship_header(h):
        return "Relationship"

    if re.search(r"\bee\b", h) and ("ssn" in h or "#" in str(header) or "not ssn" in h):
        return "Employee ID"
    if "gender" in h and ("m" in h or "f" in h):
        return "Gender"
    if "employee class" in h or "benefit class" in h:
        return "ICHRA Class"
    if "weekly" in h and "hour" in h:
        return "Hours Per Week"
    if "annual" in h and "salary" in h:
        return "Annual Salary"

    if ("home" in h or "primary address" in h) and ("zip" in h or "postal" in h):
        return "Zip Code"
    if ("work" in h or "worksite" in h or "site" in h or "work location" in h) and ("zip" in h or "postal" in h):
        return "Primary Worksite Zip Code"
    if ("home" in h or "primary address" in h) and "state" in h:
        return "State"
    if ("home" in h or "primary address" in h) and "city" in h:
        return "City"

    if "medical" in h and "election" in h and not _is_ancillary_header(h):
        return "Health Election"
    if "medical" in h and "plan" in h and "name" in h and not _is_ancillary_header(h):
        return "Current Health Plan"
    if "medical" in h and "waiver" in h:
        return "Notes"

    return None


# =============================================================================
# 3. VALUE NORMALIZATION
# =============================================================================

_RELATIONSHIP_MAP: dict[str, str] = {}


def _add_rel(canonical: str, *aliases: str) -> None:
    for alias in aliases:
        _RELATIONSHIP_MAP[_norm(alias)] = canonical


_add_rel("Employee", "employee", "ee", "emp", "subscriber", "self", "primary", "insured", "member", "worker")
_add_rel(
    "Spouse",
    "spouse", "sp", "sps", "spse", "husband", "wife", "partner", "domestic partner",
    "registered domestic partner", "common law spouse", "common law partner", "life partner",
)
_add_rel(
    "Child",
    "child", "children", "ch", "dependent", "dependant", "dep", "daughter", "son", "stepchild",
    "step child", "adopted child", "foster child", "stepson", "stepdaughter", "child of domestic partner",
)

_ELECTION_ENROLL = {
    "enroll", "enrolled", "e", "yes", "y", "participating", "active", "covered", "elect", "elected",
    "ee", "eo", "es", "ec", "ech", "eech", "ef", "fa", "fam", "family", "emp",
}
_ELECTION_WAIVE = {
    "waive", "waived", "w", "wp", "ie", "no", "n", "decline", "declined", "opt out", "optout",
    "not participating", "waiving", "waiting period", "in waiting period", "ineligible", "none",
}

_TIER_EMP_ONLY = {"employee only", "employee", "ee only", "employee only coverage", "single", "individual", "self only", "ee", "eo", "emp"}
_TIER_EMP_SPOUSE = {"employee spouse", "employee and spouse", "employee + spouse", "ee spouse", "ee + spouse", "ee sp", "employee plus spouse", "es"}
_TIER_EMP_CHILDREN = {"employee children", "employee child", "employee + children", "employee + child", "ee children", "ee child", "ee + children", "ee + child", "employee plus children", "parent child", "parent children", "ec", "ech", "eech"}
_TIER_FAMILY = {"family", "employee family", "employee + family", "ee family", "ee + family", "employee spouse children", "employee + spouse + children", "employee spouse child", "ee spouse children", "ee + spouse + children", "ef", "fa", "fam"}
_TIER_WAIVE = {"waive", "waived", "decline", "declined", "no coverage", "none", "not enrolled", "no election", "w", "wp"}

VALID_HEALTH_PLAN_TIERS = {"Employee Only", "Employee + Spouse", "Employee + Children", "Family"}
VALID_HEALTH_PLAN_TIERS_WITH_WAIVE = VALID_HEALTH_PLAN_TIERS | {"Waive"}


def normalize_relationship(value: object) -> tuple[str, Optional[str]]:
    original = str(value or "").strip()
    key = _norm(original.replace("|", " "))
    if not key:
        return "", None
    if key in _RELATIONSHIP_MAP:
        canonical = _RELATIONSHIP_MAP[key]
        note = f"Relationship normalized: '{original}' -> '{canonical}'" if canonical.lower() != original.lower() else None
        return canonical, note

    # Employer exports sometimes use compound relationship labels such as
    # Spouse-Ex, Child-Step, Step-Child, or Child / Domestic Partner. The CSA
    # import only accepts Employee, Spouse, or Child, so reduce obvious
    # spouse/child compounds to the accepted values while keeping an auto-fix
    # note in the report.
    if re.search(r"\b(spouse|spse|sps|husband|wife|partner)\b", key):
        return "Spouse", f"Relationship normalized: '{original}' -> 'Spouse'"
    if re.search(r"\b(child|children|step|stepchild|daughter|son|dependent|dependant|dep)\b", key):
        return "Child", f"Relationship normalized: '{original}' -> 'Child'"

    if FUZZY_AVAILABLE:
        match = rf_process.extractOne(key, _RELATIONSHIP_MAP.keys(), scorer=fuzz.ratio)
        if match:
            best, score, _ = match
            if score >= 86:
                canonical = _RELATIONSHIP_MAP[best]
                return canonical, f"Relationship fuzzy-matched: '{original}' -> '{canonical}' (score {score})"
    return original, None


def normalize_election(value: object) -> tuple[str, Optional[str]]:
    original = str(value or "").strip()
    key = _norm(original)
    if not key:
        return "", None
    if key in _ELECTION_ENROLL:
        canonical = "Enroll"
    elif key in _ELECTION_WAIVE:
        canonical = "Waive"
    else:
        return original, None
    note = f"Health Election normalized: '{original}' -> '{canonical}'" if original != canonical else None
    return canonical, note


def normalize_tier(value: object) -> tuple[str, Optional[str]]:
    original = str(value or "").strip()
    key = _norm(original.replace("&", " and ").replace("/", " ").replace("-", " "))
    key = key.replace(" plus ", " ").replace(" and ", " ")
    key = re.sub(r"\s+", " ", key).strip()
    if not key:
        return "", None

    compact = _compact(original)
    abbreviation_map = {
        "eo": "Employee Only",
        "ee": "Employee Only",
        "es": "Employee + Spouse",
        "ec": "Employee + Children",
        "ech": "Employee + Children",
        "eech": "Employee + Children",
        "ef": "Family",
        "fa": "Family",
        "fam": "Family",
    }
    if compact in abbreviation_map:
        canonical = abbreviation_map[compact]
        note = f"Tier normalized: '{original}' -> '{canonical}'" if original != canonical else None
        return canonical, note

    def normalized_set(options: set[str]) -> set[str]:
        return {_norm(o).replace(" and ", " ").replace(" plus ", " ") for o in options}

    if key in normalized_set(_TIER_WAIVE) or any(x in compact for x in ["waive", "decline", "nocoverage"]):
        return "Waive", f"Tier normalized: '{original}' -> 'Waive'" if original != "Waive" else None
    if key in normalized_set(_TIER_FAMILY) or "family" in compact or ("spouse" in compact and ("child" in compact or "children" in compact)):
        return "Family", f"Tier normalized: '{original}' -> 'Family'" if original != "Family" else None
    if key in normalized_set(_TIER_EMP_SPOUSE) or ("spouse" in compact and "child" not in compact and "children" not in compact):
        return "Employee + Spouse", f"Tier normalized: '{original}' -> 'Employee + Spouse'" if original != "Employee + Spouse" else None
    if key in normalized_set(_TIER_EMP_CHILDREN) or "child" in compact or "children" in compact:
        return "Employee + Children", f"Tier normalized: '{original}' -> 'Employee + Children'" if original != "Employee + Children" else None
    if key in normalized_set(_TIER_EMP_ONLY) or compact in {"employeeonly", "selfonly", "individual"}:
        return "Employee Only", f"Tier normalized: '{original}' -> 'Employee Only'" if original != "Employee Only" else None
    return original, None


def derive_tier_from_coverage_code(value: object) -> tuple[str, Optional[str]]:
    original = str(value or "").strip()
    compact = _compact(original)
    if not compact:
        return "", None
    code_map = {
        "ee": "Employee Only",
        "eo": "Employee Only",
        "es": "Employee + Spouse",
        "ec": "Employee + Children",
        "ech": "Employee + Children",
        "eech": "Employee + Children",
        "ef": "Family",
        "fa": "Family",
        "fam": "Family",
    }
    if compact in code_map:
        canonical = code_map[compact]
        return canonical, f"Current Health Plan Tier derived from coverage code '{original}' -> '{canonical}'"

    tier, note = normalize_tier(original)
    if tier in {"Employee Only", "Employee + Spouse", "Employee + Children", "Family"}:
        # Avoid treating a plain "E" / "Enroll" health election as Employee Only.
        if compact not in {"e", "enroll", "enrolled", "elect", "elected"}:
            return tier, note or f"Current Health Plan Tier derived from '{original}' -> '{tier}'"
    return "", None


def dependent_election_from_tier(employee_election: str, tier: str, relationship: str) -> str:
    if employee_election == "Waive" or tier == "Waive":
        return "Waive"
    if relationship == "Employee":
        return employee_election or "Enroll"
    if tier == "Employee Only":
        return "Waive"
    if tier == "Employee + Spouse":
        return "Enroll" if relationship == "Spouse" else "Waive"
    if tier == "Employee + Children":
        return "Enroll" if relationship == "Child" else "Waive"
    if tier == "Family":
        return "Enroll"
    return employee_election or "Enroll"


# =============================================================================
# 4. ZIP CODE LOOKUP
# =============================================================================

@st.cache_data(show_spinner=False)
def build_zip_cache() -> dict[str, dict[str, str]]:
    if not ZIPCODES_AVAILABLE:
        return {}
    cache: dict[str, dict[str, str]] = {}
    for entry in _zc.list_all():
        zip_code = str(entry.get("zip_code", "")).strip()
        state = str(entry.get("state", "")).strip().upper()
        zip_type = str(entry.get("zip_code_type", "")).strip().upper()
        active = bool(entry.get("active", True))
        if not zip_code or zip_code in KNOWN_INVALID_OR_UNSUPPORTED_ZIPS:
            continue
        if state not in SUPPORTED_ZIP_STATES:
            continue
        if not active:
            continue
        if zip_type and zip_type != "STANDARD":
            continue
        county_raw = str(entry.get("county") or "")
        county = county_raw.replace(" County", "").strip()
        cache.setdefault(zip_code, {"abbr": state, "county": county})
    return cache


def _pad_zip(value: object) -> str:
    v = str(value or "").strip().upper()
    if not v:
        return ""
    if re.fullmatch(r"\d+\.0", v):
        v = v[:-2]
    if re.search(r"[A-Z]", v):
        return v
    compact = re.sub(r"[\s-]", "", v)
    if re.fullmatch(r"\d{9}", compact):
        return compact[:5]
    if re.fullmatch(r"\d{5}", compact):
        return compact
    if re.fullmatch(r"\d{1,4}", compact):
        return compact.zfill(5)
    return v


def lookup_zip(value: object, cache: dict[str, dict[str, str]]) -> Optional[dict[str, str]]:
    return cache.get(_pad_zip(value)[:5])


# =============================================================================
# 5. FILE INGESTION WITH HEADER ROW DETECTION
# =============================================================================

CANDIDATE_ENCODINGS = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]


def _detect_delimiter(text: str) -> str:
    nonblank_lines = [line for line in text.splitlines() if line.strip()]
    sample_lines = nonblank_lines[:10]
    delimiter_scores = {
        ",": sum(line.count(",") for line in sample_lines),
        "\t": sum(line.count("\t") for line in sample_lines),
        "|": sum(line.count("|") for line in sample_lines),
        ";": sum(line.count(";") for line in sample_lines),
    }
    delimiter = max(delimiter_scores, key=delimiter_scores.get)
    return delimiter if delimiter_scores[delimiter] > 0 else ","


def _score_header_row(values: list[object]) -> int:
    score = 0
    seen: set[str] = set()
    for value in values:
        canonical = canonical_from_header(value)
        if canonical:
            score += 3 if canonical in REQUIRED_FIELDS else 2
            seen.add(canonical)
        elif _is_actual_dependent_header(value):
            score += 2
    if {"First Name", "Last Name", "DOB"}.issubset(seen):
        score += 4
    if "Relationship" in seen:
        score += 3
    if "Zip Code" in seen:
        score += 2
    if "Primary Worksite Zip Code" in seen:
        score += 2
    return score


def _dedupe_headers(headers: list[object]) -> list[str]:
    output: list[str] = []
    counts: dict[str, int] = {}
    for idx, header in enumerate(headers):
        name = str(header or "").strip()
        if not name or name.lower().startswith("unnamed"):
            name = f"Unnamed Column {idx + 1}"
        if name in counts:
            counts[name] += 1
            name = f"{name}.{counts[name]}"
        else:
            counts[name] = 0
        output.append(name)
    return output


def _row_values_are_blank(values: list[object]) -> bool:
    return all(str(value or "").strip() == "" for value in values)


def _row_has_footer_marker(values: list[object]) -> bool:
    row_text = " ".join(str(value or "").strip() for value in values if str(value or "").strip())
    normalized = _norm(row_text)
    if not normalized:
        return False

    footer_patterns = [
        r"\bcoverage key\b",
        r"\bcoverage legend\b",
        r"\belection key\b",
        r"\bbenefit key\b",
        r"\binstructions?\b",
        r"\bnotes?\b.*\bmust include\b",
        r"\bmust include\b.*\b(date of birth|dob|gender)\b",
        r"\bee employee only\b",
        r"\besp employee spouse\b",
        r"\bech employee child",
        r"\bfam employee family\b",
        r"\btotal employees?\b",
        r"\btotal lives?\b",
    ]
    return any(re.search(pattern, normalized) for pattern in footer_patterns)


def _looks_like_date_value(value: object) -> bool:
    raw = str(value or "").strip()
    if not raw:
        return False
    if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{2,4}", raw):
        return True
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}(\s+\d{1,2}:\d{2}:\d{2})?", raw):
        return True
    if re.fullmatch(r"\d{5}", raw):
        # Could be either an Excel serial date or a ZIP. It still counts as a
        # census signal when it appears under a DOB/birth-date column.
        return True
    try:
        parsed = pd.to_datetime(raw, errors="coerce")
        return pd.notna(parsed)
    except Exception:
        return False


def _looks_like_zip_value(value: object) -> bool:
    raw = str(value or "").strip()
    if not raw:
        return False
    return bool(re.fullmatch(r"\d{5}(-?\d{4})?", raw) or re.fullmatch(r"\d{1,4}", raw))


def _looks_like_coverage_code(value: object) -> bool:
    compact = _compact(value)
    return compact in {"ee", "eo", "es", "esp", "ec", "ech", "eech", "ef", "fa", "fam", "family", "waive", "w"}


def _trim_census_body_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Remove title/footer material below the actual census body.

    Header detection already handles extra rows above the census. This pass
    handles the opposite problem: notes, coverage legends, and instruction blocks
    below the census. It preserves the data body, stops after a clear blank gap,
    and stops immediately when a footer marker such as "Coverage Key" appears.
    """
    if df.empty:
        return df

    headers = list(df.columns)
    canonical_by_col = {col: canonical_from_header(col) for col in headers}
    kept_indices: list[int] = []
    data_started = False
    consecutive_blank_after_data = 0

    def row_has_census_signal(row: pd.Series) -> bool:
        first_last_name_present = False
        relationship_signal = False
        date_signal = False
        zip_signal = False
        coverage_signal = False
        dependent_signal = False

        for col in headers:
            value = str(row.get(col, "") or "").strip()
            if not value:
                continue
            canonical = canonical_by_col.get(col)
            header_norm = _norm(col)

            if canonical in {"First Name", "Last Name"}:
                first_last_name_present = True
            if canonical == "Relationship":
                rel, _ = normalize_relationship(value)
                relationship_signal = rel in {"Employee", "Spouse", "Child"}
            if canonical == "DOB" or "dob" in header_norm or "birth" in header_norm:
                date_signal = _looks_like_date_value(value)
            if canonical in {"Zip Code", "Primary Worksite Zip Code"} or "zip" in header_norm:
                zip_signal = _looks_like_zip_value(value)
            if canonical == "Health Election" or header_norm in {"medical", "medical coverage", "coverage"}:
                coverage_signal = _looks_like_coverage_code(value) or normalize_election(value)[0] in {"Enroll", "Waive"}
            if _is_actual_dependent_header(col):
                dependent_signal = True

        strong_signal = relationship_signal or date_signal or zip_signal or coverage_signal or dependent_signal
        return strong_signal or (first_last_name_present and strong_signal)

    for idx, row in df.iterrows():
        values = row.tolist()
        is_blank = _row_values_are_blank(values)

        if is_blank:
            if data_started:
                consecutive_blank_after_data += 1
                if consecutive_blank_after_data >= 2:
                    break
            continue

        consecutive_blank_after_data = 0

        if data_started and _row_has_footer_marker(values):
            break

        if row_has_census_signal(row):
            kept_indices.append(idx)
            data_started = True
            continue

        # A non-census row after data has started is usually a footer, legend,
        # or note block. Stop instead of pushing it downstream as an employee.
        if data_started:
            break

    if not kept_indices:
        return df.loc[df.apply(lambda row: any(str(v).strip() for v in row), axis=1)].reset_index(drop=True)
    return df.loc[kept_indices].reset_index(drop=True)


def _finalize_raw_dataframe(raw: pd.DataFrame) -> pd.DataFrame:
    raw = raw.fillna("").astype(str)
    nonblank_mask = raw.apply(lambda row: any(str(v).strip() for v in row), axis=1)
    if not nonblank_mask.any():
        return pd.DataFrame()

    first_nonblank = int(nonblank_mask[nonblank_mask].index[0])
    last_nonblank = int(nonblank_mask[nonblank_mask].index[-1])
    raw = raw.iloc[first_nonblank:last_nonblank + 1].reset_index(drop=True)

    max_scan = min(len(raw), 30)
    scored = [(idx, _score_header_row(raw.iloc[idx].tolist())) for idx in range(max_scan)]
    best_idx, best_score = max(scored, key=lambda item: item[1])
    header_idx = best_idx if best_score >= 5 else 0

    headers = _dedupe_headers(raw.iloc[header_idx].tolist())
    df = raw.iloc[header_idx + 1:].reset_index(drop=True).copy()
    df.columns = headers
    df = _trim_census_body_rows(df)
    return df.fillna("").astype(str).reset_index(drop=True)

def _read_csv_text(text: str) -> pd.DataFrame:
    delimiter = _detect_delimiter(text)
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows = [list(row) for row in reader]
    if not rows or not any(any(str(cell).strip() for cell in row) for row in rows):
        return pd.DataFrame()
    max_width = max(len(row) for row in rows)
    normalized_rows = [row + [""] * (max_width - len(row)) for row in rows]
    raw = pd.DataFrame(normalized_rows, dtype=str)
    return _finalize_raw_dataframe(raw)


def _read_csv_with_fallback_encoding(file_obj) -> tuple[pd.DataFrame, str]:
    raw_bytes = file_obj.read()
    if isinstance(raw_bytes, str):
        return _read_csv_text(raw_bytes), "text"

    last_exc: Optional[Exception] = None
    for encoding in CANDIDATE_ENCODINGS:
        try:
            text = raw_bytes.decode(encoding)
            return _read_csv_text(text), encoding
        except Exception as exc:
            last_exc = exc
            continue

    try:
        text = raw_bytes.decode("latin-1", errors="replace")
        return _read_csv_text(text), "latin-1 with replacement"
    except Exception:
        raise last_exc if last_exc else RuntimeError("Unable to decode uploaded file")


def _read_excel_with_fallback(uploaded) -> pd.DataFrame:
    raw_bytes = uploaded.read()
    if isinstance(raw_bytes, str):
        raw_bytes = raw_bytes.encode("utf-8")

    last_exc: Optional[Exception] = None
    for engine in [None, "openpyxl", "xlrd"]:
        try:
            bio = io.BytesIO(raw_bytes)
            kwargs = {"dtype": str, "keep_default_na": False, "header": None}
            if engine is not None:
                kwargs["engine"] = engine
            raw = pd.read_excel(bio, **kwargs)
            return _finalize_raw_dataframe(raw)
        except ImportError as exc:
            last_exc = exc
            continue
        except ValueError as exc:
            last_exc = exc
            continue
        except Exception as exc:
            last_exc = exc
            continue

    msg = str(last_exc or "")
    compound_file_signature = raw_bytes.startswith(b"\xd0\xcf\x11\xe0")
    if compound_file_signature or "encrypted" in msg.lower() or "ole2 compound" in msg.lower():
        raise ValueError(
            "Could not read the uploaded Excel file. It appears to be encrypted, password-protected, "
            "or not saved as a standard Excel workbook. Save a non-password-protected .xlsx or CSV copy "
            "and upload that file."
        ) from last_exc
    raise ValueError(f"Could not read the uploaded Excel file: {last_exc}") from last_exc


def read_uploaded_file(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        df, _ = _read_csv_with_fallback_encoding(uploaded)
        return df
    if name.endswith((".xlsx", ".xls")):
        return _read_excel_with_fallback(uploaded)
    raise ValueError(f"Unsupported file type: {uploaded.name}")


# =============================================================================
# 6. COLUMN MAPPING
# =============================================================================

FUZZY_THRESHOLD = 88


def map_columns(df: pd.DataFrame) -> tuple[dict[int, int], list[str], list[str]]:
    raw_headers = [str(c) for c in df.columns]
    header_canonicals = [canonical_from_header(h) for h in raw_headers]
    mapping: dict[int, int] = {}
    used_df_cols: set[int] = set()
    mapping_notes: list[str] = []

    for fi, field_def in enumerate(FIELDS):
        field_name = field_def["name"]
        for ci, canonical in enumerate(header_canonicals):
            if ci in used_df_cols:
                continue
            if canonical == field_name:
                mapping[fi] = ci
                used_df_cols.add(ci)
                break

    if FUZZY_AVAILABLE:
        for fi, field_def in enumerate(FIELDS):
            if fi in mapping:
                continue
            field_name = field_def["name"]
            candidates = [ci for ci in range(len(raw_headers)) if ci not in used_df_cols]
            best_ci: Optional[int] = None
            best_score = 0
            for ci in candidates:
                hn = _norm(raw_headers[ci])
                # Avoid fuzzy-mapping ancillary benefit columns to health-plan fields.
                if field_name.startswith("Current Health") and _is_ancillary_header(hn):
                    continue
                score = max(
                    fuzz.ratio(field_name.lower(), hn),
                    fuzz.token_sort_ratio(field_name.lower(), hn),
                    fuzz.partial_ratio(field_name.lower(), hn),
                )
                if score > best_score:
                    best_score = score
                    best_ci = ci
            if best_ci is not None and best_score >= FUZZY_THRESHOLD:
                mapping[fi] = best_ci
                used_df_cols.add(best_ci)
                mapping_notes.append(f"Fuzzy mapped '{raw_headers[best_ci]}' -> '{field_name}' (score {best_score})")

    unrecognized = [raw_headers[ci] for ci in range(len(raw_headers)) if ci not in used_df_cols]
    return mapping, unrecognized, mapping_notes


# =============================================================================
# 7. HORIZONTAL CENSUS DETECTION AND CONVERSION
# =============================================================================

@dataclass
class HorizontalOptions:
    auto_placeholder_names: bool = True
    infer_dependents_enrolled_when_tier_missing: bool = True
    inherit_worksite_zip_to_dependents: bool = True
    exclude_interns_and_contractors: bool = True


@dataclass
class HorizontalColumnMap:
    employee_first: Optional[str] = None
    employee_last: Optional[str] = None
    employee_id: Optional[str] = None
    employee_dob: Optional[str] = None
    employee_gender: Optional[str] = None
    employee_email: Optional[str] = None
    address1: Optional[str] = None
    address2: Optional[str] = None
    city: Optional[str] = None
    state: Optional[str] = None
    zip_code: Optional[str] = None
    county: Optional[str] = None
    worksite_zip: Optional[str] = None
    worksite_county: Optional[str] = None
    ichra_class: Optional[str] = None
    worker_category: Optional[str] = None
    position_status: Optional[str] = None
    health_election: Optional[str] = None
    health_waive_reason: Optional[str] = None
    plan_vendor: Optional[str] = None
    plan_name: Optional[str] = None
    plan_tier: Optional[str] = None
    oop_single: Optional[str] = None
    oop_family: Optional[str] = None
    ded_single: Optional[str] = None
    ded_family: Optional[str] = None
    er_cost: Optional[str] = None
    ee_cost: Optional[str] = None
    annual_salary: Optional[str] = None
    hourly_rate: Optional[str] = None
    hours_per_week: Optional[str] = None
    notes: Optional[str] = None


@dataclass
class DepSlot:
    slot: str
    number: int = 9999
    relationship_col: Optional[str] = None
    dob_col: Optional[str] = None
    first_col: Optional[str] = None
    last_col: Optional[str] = None
    default_relationship: str = ""


def _pick_col(
    df: pd.DataFrame,
    aliases: list[str],
    *,
    health_only: bool = False,
    dependent_only: bool = False,
    allow_contains: bool = True,
    allow_fuzzy: bool = True,
) -> Optional[str]:
    headers = list(df.columns)
    norm_to_header: dict[str, str] = {_norm(h): h for h in headers}
    alias_norms = [_norm(a) for a in aliases]

    def allowed(header: str) -> bool:
        if health_only and not _looks_health_header(header):
            return False
        if dependent_only and not _is_actual_dependent_header(header):
            return False
        return True

    def contains_match(header_norm: str, alias_norm: str) -> bool:
        if not alias_norm:
            return False
        if header_norm == alias_norm:
            return True
        if len(alias_norm) >= 4 and alias_norm in header_norm:
            return True

        # Avoid mapping generic one-word headers like "Medical" into every
        # medical plan/cost field just because the alias contains the word
        # "medical". Reverse containment is only safe for descriptive headers.
        generic_headers = {"medical", "health", "plan", "coverage", "carrier", "name", "cost", "tier"}
        if header_norm in generic_headers:
            return False
        if len(header_norm) >= 8 and len(header_norm.split()) >= 2 and header_norm in alias_norm:
            return True
        return False

    for alias in alias_norms:
        header = norm_to_header.get(alias)
        if header and allowed(header):
            return header

    if allow_contains:
        for header in headers:
            hn = _norm(header)
            if not allowed(header):
                continue
            if any(contains_match(hn, alias) for alias in alias_norms):
                return header

    if allow_fuzzy and FUZZY_AVAILABLE:
        searchable = [_norm(h) for h in headers if allowed(h)]
        for alias in alias_norms:
            match = rf_process.extractOne(alias, searchable, scorer=fuzz.token_sort_ratio)
            if match:
                best_norm, score, _ = match
                if score >= 95:
                    return norm_to_header.get(best_norm)
    return None

def build_horizontal_column_map(df: pd.DataFrame) -> HorizontalColumnMap:
    cm = HorizontalColumnMap()
    cm.employee_first = _pick_col(df, ["First Name", "Employee First Name", "EE First Name", "First", "Given Name"])
    cm.employee_last = _pick_col(df, ["Last Name", "Employee Last Name", "EE Last Name", "Last", "Surname", "Family Name"])
    cm.employee_id = _pick_col(df, ["Employee ID", "Employee Number", "Employee #", "EE #", "Worker ID", "Person Number"], allow_contains=False, allow_fuzzy=False)
    cm.employee_dob = _pick_col(df, ["BIRTH DATE", "DOB", "Employee DOB", "Employee Date of Birth", "Date of Birth", "Birthdate"])
    cm.employee_gender = _pick_col(df, ["Gender", "Sex"])
    cm.employee_email = _pick_col(df, ["Email", "Email Address", "Work Email", "Personal Email"])
    cm.address1 = _pick_col(df, ["PRIMARY ADDRESS LINE 1", "Address Line 1", "Street Address", "Home Address Line 1"])
    cm.address2 = _pick_col(df, ["PRIMARY ADDRESS LINE 2", "Address Line 2", "Home Address Line 2"])
    cm.city = _pick_col(df, ["PRIMARY ADDRESS - CITY", "Primary Address City", "City", "Home City"])
    cm.state = _pick_col(df, ["PRIMARY ADDRESS - STATE / TERRITORY", "Primary Address State", "State", "Home State"])
    cm.zip_code = _pick_col(df, ["PRIMARY ADDRESS - ZIP CODE", "Primary Address Zip Code", "Zip Code", "Zip", "Home Zip"])
    cm.county = _pick_col(df, ["County", "Home County", "Primary Address County"])
    cm.worksite_zip = _pick_col(df, ["WORKSITE ZIP CODE", "Primary Worksite Zip Code", "Worksite Zip", "Work Zip", "Work Location Zip"])
    cm.worksite_county = _pick_col(df, ["Primary Worksite County", "Worksite County", "Work County", "Work Location County"])
    cm.ichra_class = _pick_col(df, ["ICHRA Class", "Benefit Class", "Employee Class", "Class"])
    cm.worker_category = _pick_col(df, ["WORKER CATEGORY", "Worker Category", "Employee Type", "Employment Type"])
    cm.position_status = _pick_col(df, ["POSITION STATUS", "Employment Status", "Status"])
    cm.health_election = _pick_col(df, ["Health Election", "Medical", "Medical Coverage", "Medical Coverage Code", "Medical Election", "Medical Coverage Election", "Coverage Election"], health_only=True)
    cm.health_waive_reason = _pick_col(df, ["MEDICAL WAIVE REASON", "Health Waive Reason", "Waive Reason"], health_only=True)
    cm.plan_vendor = _pick_col(df, ["Current Health Plan Vendor", "Health Carrier", "Medical Carrier", "Carrier", "Medical Vendor"], health_only=True)
    cm.plan_name = _pick_col(df, ["Product Name", "Current Plan Name", "Current Product Name", "Medical Product Name", "Health Product Name"], health_only=True, allow_fuzzy=False) or _pick_col(df, ["MEDICAL PLAN NAME", "Medical Plan Name", "Health Plan Name", "Current Health Plan", "Plan Name"], health_only=True)
    cm.plan_tier = _pick_col(df, ["Current Health Plan Tier", "Medical Plan Tier", "Medical Coverage Tier", "Coverage Tier", "Coverage Level", "Benefit Tier", "Election Tier"], health_only=True)
    cm.oop_single = _pick_col(df, ["Current Health Plan OOP (single)", "Medical OOP Single", "OOP Single"], health_only=True)
    cm.oop_family = _pick_col(df, ["Current Health Plan OOP (family)", "Medical OOP Family", "OOP Family"], health_only=True)
    cm.ded_single = _pick_col(df, ["Current Health Plan Deductible (single)", "Medical Deductible Single", "Deductible Single"], health_only=True)
    cm.ded_family = _pick_col(df, ["Current Health Plan Deductible (family)", "Medical Deductible Family", "Deductible Family"], health_only=True)
    cm.er_cost = _pick_col(df, ["Current Health Plan ER Cost", "Medical ER Cost", "Employer Cost", "Pending Employer Cost", "Employer Contribution"], health_only=True, allow_fuzzy=False)
    cm.ee_cost = _pick_col(df, ["Current Health Plan EE Cost", "Medical EE Cost", "Employee Cost", "Pending Employee Cost", "Employee Contribution"], health_only=True, allow_fuzzy=False)
    cm.annual_salary = _pick_col(df, ["ANNUAL SALARY", "Annual Salary", "Salary", "Base Salary"])
    cm.hourly_rate = _pick_col(df, ["Hourly Rate", "Hourly Wage"])
    cm.hours_per_week = _pick_col(df, ["Hours Per Week", "Weekly Hours", "Weekly Hours Worked"])
    cm.notes = _pick_col(df, ["Notes", "Comments", "Remarks"])
    return cm


def find_dependent_slots(df: pd.DataFrame) -> tuple[list[DepSlot], bool]:
    row_rel = _pick_col(df, ["Dependent Relationship", "Dependent Relation", "Dependent Type", "Dep Relationship", "Dep Type"], dependent_only=True, allow_fuzzy=False)
    row_dob = _pick_col(df, ["Dependent DOB", "Dependent Date of Birth", "Dependent Birth Date", "Dep DOB", "Dep Birth Date"], dependent_only=True, allow_fuzzy=False)
    row_first = _pick_col(df, ["Dependent First Name", "Dep First Name", "Dependent First"], dependent_only=True, allow_fuzzy=False)
    row_last = _pick_col(df, ["Dependent Last Name", "Dep Last Name", "Dependent Last"], dependent_only=True, allow_fuzzy=False)
    if row_rel or row_dob or row_first or row_last:
        return [DepSlot("rowwise", relationship_col=row_rel, dob_col=row_dob, first_col=row_first, last_col=row_last, number=1)], True

    headers = list(df.columns)
    header_positions = {header: idx for idx, header in enumerate(headers)}
    slots: dict[str, DepSlot] = {}

    def get_slot(key: str, number: int, default_relationship: str = "") -> DepSlot:
        if key not in slots:
            slots[key] = DepSlot(slot=key, number=number, default_relationship=default_relationship)
        elif default_relationship and not slots[key].default_relationship:
            slots[key].default_relationship = default_relationship
        return slots[key]

    for header in headers:
        hn = _norm(header)
        if not _is_actual_dependent_header(header):
            continue
        if _is_ancillary_header(header):
            continue
        number_match = re.search(r"(\d+)", hn)
        number = int(number_match.group(1)) if number_match else 999

        if "spouse" in hn or "domestic partner" in hn:
            slot_number = number if number != 999 else 1
            slot = get_slot(f"spouse_{slot_number}", 100 + slot_number, "Spouse")
        else:
            is_child_header = bool(re.search(r"\b(child|children|ch)\b", hn))
            is_numbered_dependent_header = bool(
                re.search(r"\b(dependent|dependant|dep)\b", hn)
                and number != 999
            )
            default_rel = "Child" if (is_child_header or is_numbered_dependent_header) else ""
            slot = get_slot(f"dep_{number}", 200 + number, default_rel)

        if any(x in hn for x in ["dob", "date of birth", "birth date", "birthdate", "birth"]):
            slot.dob_col = header
        elif "first" in hn and "name" in hn:
            slot.first_col = header
        elif "last" in hn and "name" in hn:
            slot.last_col = header
        elif any(x in hn for x in ["relationship", "relation", "type"]):
            slot.relationship_col = header

    # Many employer templates have a generic "Name" column immediately before
    # each dependent DOB column, for example: Name, SPOUSE DOB, SPOUSE GENDER,
    # Name, CH #1 DOB. Since duplicate headers are deduped as Name.1, Name.2,
    # attach those generic name columns to the nearest dependent slot on their
    # right.
    generic_name_headers = [
        header for header in headers
        if re.fullmatch(r"name( \d+)?", _norm(header))
    ]
    for slot in slots.values():
        if slot.first_col:
            continue
        anchor_col = slot.dob_col or slot.relationship_col or slot.last_col
        if not anchor_col or anchor_col not in header_positions:
            continue
        anchor_idx = header_positions[anchor_col]
        candidate_headers = [
            header for header in generic_name_headers
            if header_positions[header] < anchor_idx and anchor_idx - header_positions[header] <= 2
        ]
        if candidate_headers:
            slot.first_col = sorted(candidate_headers, key=lambda h: anchor_idx - header_positions[h])[0]

    result = [slot for slot in slots.values() if slot.relationship_col or slot.dob_col or slot.first_col or slot.last_col]
    result.sort(key=lambda s: (0 if s.default_relationship == "Spouse" else 1, s.number, s.slot))
    return result, False

def detect_horizontal_census(df: pd.DataFrame) -> tuple[bool, str]:
    headers = list(df.columns)
    actual_dependent_headers = [h for h in headers if _is_actual_dependent_header(h)]
    canonical_headers = {canonical_from_header(h) for h in headers if canonical_from_header(h)}
    vertical_score = len(canonical_headers.intersection({"First Name", "Last Name", "Relationship", "DOB", "Zip Code", "Primary Worksite Zip Code", "Health Election"}))

    if actual_dependent_headers:
        return True, "Actual dependent-specific columns detected"
    if vertical_score >= 4 and "Relationship" in canonical_headers and "DOB" in canonical_headers:
        return False, "Standard vertical format detected"
    return False, "Standard vertical format detected"


def _cell(row: pd.Series, col: Optional[str]) -> str:
    if not col or col not in row.index:
        return ""
    return str(row[col] or "").strip()


def _first_nonblank(rows: pd.DataFrame, col: Optional[str]) -> str:
    if not col or col not in rows.columns:
        return ""
    for value in rows[col].tolist():
        if str(value).strip():
            return str(value).strip()
    return ""


def _derive_ichra_class(row: pd.Series, cm: HorizontalColumnMap) -> str:
    direct = _cell(row, cm.ichra_class)
    if direct:
        return direct
    worker_category = _cell(row, cm.worker_category)
    state = _cell(row, cm.state).upper()
    if not worker_category:
        return "NY Full Time Benefits Eligible" if state in {"NY", "VT"} else "Full Time Benefits Eligible"
    wc = _norm(worker_category)
    if "intern" in wc or "contractor" in wc:
        return ""
    if "part" in wc:
        return "NY Part Time Benefits Eligible" if state in {"NY", "VT"} else "Part Time Benefits Eligible"
    return "NY Full Time Benefits Eligible" if state in {"NY", "VT"} else "Full Time Benefits Eligible"


def _should_exclude_employee(row: pd.Series, cm: HorizontalColumnMap, options: HorizontalOptions) -> bool:
    if not options.exclude_interns_and_contractors:
        return False
    wc = _norm(_cell(row, cm.worker_category))
    return "intern" in wc or "contractor" in wc


def _employee_group_key(row: pd.Series, cm: HorizontalColumnMap, row_index: int) -> tuple[Any, ...]:
    employee_id = _cell(row, cm.employee_id)
    if employee_id:
        return ("employee_id", employee_id)
    first = _cell(row, cm.employee_first)
    last = _cell(row, cm.employee_last)
    dob = _cell(row, cm.employee_dob)
    if first or last:
        return ("name_dob", first, last, dob, _cell(row, cm.zip_code))
    stable_values = [
        _cell(row, cm.employee_dob), _cell(row, cm.address1), _cell(row, cm.city), _cell(row, cm.state),
        _cell(row, cm.zip_code), _cell(row, cm.worksite_zip), _cell(row, cm.worker_category),
        _cell(row, cm.position_status), _cell(row, cm.annual_salary), _cell(row, cm.plan_name),
        _cell(row, cm.plan_tier), _cell(row, cm.health_waive_reason),
    ]
    if any(stable_values):
        return ("anonymous", *stable_values)
    return ("row", row_index)


def _sort_dependents(dependents: list[dict[str, str]]) -> list[dict[str, str]]:
    def sort_key(dep: dict[str, str]) -> tuple[int, int]:
        relationship = dep.get("Relationship", "")
        relationship_order = 0 if relationship == "Spouse" else 1 if relationship == "Child" else 2
        number_match = re.search(r"\d+", str(dep.get("_dep_number", "9999")))
        number = int(number_match.group(0)) if number_match else 9999
        return relationship_order, number
    return sorted(dependents, key=sort_key)


def _infer_tier_from_dependents(dependents: list[dict[str, str]]) -> str:
    has_spouse = any(dep.get("Relationship") == "Spouse" for dep in dependents)
    has_child = any(dep.get("Relationship") == "Child" for dep in dependents)
    if has_spouse and has_child:
        return "Family"
    if has_spouse:
        return "Employee + Spouse"
    if has_child:
        return "Employee + Children"
    return "Employee Only"


def convert_horizontal_to_vertical(df: pd.DataFrame, options: HorizontalOptions) -> dict[str, Any]:
    df = df.fillna("").astype(str).reset_index(drop=True)
    cm = build_horizontal_column_map(df)
    dep_slots, rowwise = find_dependent_slots(df)
    notes: list[dict[str, str]] = []
    output_rows: list[dict[str, str]] = []
    placeholder_counter = 1

    if not dep_slots:
        notes.append({"Kind": "Warning", "Issue": "Horizontal mode was selected, but no dependent DOB/relationship columns were detected."})

    groups: list[tuple[tuple[Any, ...], list[int]]] = []
    seen: dict[tuple[Any, ...], int] = {}
    for idx, row in df.iterrows():
        if all(_is_blank(v) for v in row):
            continue
        if _should_exclude_employee(row, cm, options):
            notes.append({"Kind": "Warning", "Issue": f"Input row {idx + 2} skipped because Worker Category appears to be intern/contractor."})
            continue
        key = _employee_group_key(row, cm, int(idx)) if rowwise else ("single_row", int(idx))
        if key not in seen:
            seen[key] = len(groups)
            groups.append((key, []))
        groups[seen[key]][1].append(int(idx))

    def base_row() -> dict[str, str]:
        return {column: "" for column in OUTPUT_COLUMNS}

    def apply_placeholder(row_out: dict[str, str], relationship: str) -> None:
        nonlocal placeholder_counter
        if row_out["First Name"] or row_out["Last Name"]:
            return
        if options.auto_placeholder_names:
            row_out["First Name"] = relationship or "Member"
            row_out["Last Name"] = str(placeholder_counter)
            placeholder_counter += 1

    for _, idxs in groups:
        rows = df.loc[idxs]
        first_row = rows.iloc[0]
        dependents: list[dict[str, str]] = []

        if rowwise and dep_slots:
            slot = dep_slots[0]
            dep_num_col = _pick_col(df, ["Dependent Number", "Dependent #", "Dep Number", "Dep #"], dependent_only=True, allow_fuzzy=False)
            for idx in idxs:
                src = df.loc[idx]
                rel_raw = _cell(src, slot.relationship_col)
                dob = _cell(src, slot.dob_col)
                first = _cell(src, slot.first_col)
                last = _cell(src, slot.last_col)
                dep_num = _cell(src, dep_num_col) or str(len(dependents) + 1)
                if not any([rel_raw, dob, first, last]):
                    continue
                rel_parts = [x.strip() for x in re.split(r"\s*[|;/]\s*", rel_raw) if x.strip()] or [rel_raw]
                dob_parts = [x.strip() for x in re.split(r"\s*[|;]\s*", dob) if x.strip()] or [dob]
                first_parts = [x.strip() for x in re.split(r"\s*[|;]\s*", first) if x.strip()] or [first]
                last_parts = [x.strip() for x in re.split(r"\s*[|;]\s*", last) if x.strip()] or [last]
                part_count = max(len(rel_parts), len(dob_parts), len(first_parts), len(last_parts), 1)
                if part_count > 1:
                    notes.append({"Kind": "Auto-fix", "Issue": f"Input row {idx + 2}: split one dependent cell into {part_count} dependent rows."})
                for part_i in range(part_count):
                    rel_piece = rel_parts[part_i] if part_i < len(rel_parts) else (rel_parts[-1] if rel_parts else "")
                    dob_piece = dob_parts[part_i] if part_i < len(dob_parts) else ""
                    first_piece = first_parts[part_i] if part_i < len(first_parts) else ""
                    last_piece = last_parts[part_i] if part_i < len(last_parts) else ""
                    relationship, note = normalize_relationship(rel_piece)
                    if note:
                        notes.append({"Kind": "Auto-fix", "Issue": f"Input row {idx + 2}: {note}"})
                    if relationship == "Employee":
                        continue
                    dependents.append({
                        "Relationship": relationship,
                        "DOB": dob_piece,
                        "First Name": first_piece,
                        "Last Name": last_piece,
                        "_dep_number": f"{dep_num}.{part_i + 1}" if part_count > 1 else dep_num,
                        "_source_row": str(idx + 2),
                    })
        else:
            for slot in dep_slots:
                explicit_rel_raw = _cell(first_row, slot.relationship_col)
                dob = _cell(first_row, slot.dob_col)
                first = _cell(first_row, slot.first_col)
                last = _cell(first_row, slot.last_col)
                if not any([explicit_rel_raw, dob, first, last]):
                    continue
                rel_raw = explicit_rel_raw or slot.default_relationship
                relationship, note = normalize_relationship(rel_raw)
                if note:
                    notes.append({"Kind": "Auto-fix", "Issue": f"Input row {idxs[0] + 2}: {note}"})
                if relationship == "Employee":
                    continue
                dependents.append({
                    "Relationship": relationship,
                    "DOB": dob,
                    "First Name": first,
                    "Last Name": last,
                    "_dep_number": str(slot.number),
                    "_source_row": str(idxs[0] + 2),
                })

        dependents = _sort_dependents(dependents)

        raw_election = _first_nonblank(rows, cm.health_election)
        raw_waive_reason = _first_nonblank(rows, cm.health_waive_reason)
        raw_plan_name = _first_nonblank(rows, cm.plan_name)
        raw_tier = _first_nonblank(rows, cm.plan_tier)
        tier, tier_note = normalize_tier(raw_tier)
        if not tier:
            tier, tier_note = derive_tier_from_coverage_code(raw_election)
        if tier_note:
            notes.append({"Kind": "Auto-fix", "Issue": tier_note})

        employee_election, election_note = normalize_election(raw_election)
        if election_note:
            notes.append({"Kind": "Auto-fix", "Issue": election_note})
        if not employee_election:
            if tier == "Waive" or raw_waive_reason:
                employee_election = "Waive"
            elif raw_plan_name or (tier and tier != "Waive"):
                employee_election = "Enroll"
            else:
                employee_election = "Waive"
                notes.append({"Kind": "Warning", "Issue": f"Input row {idxs[0] + 2}: no health election, plan, or tier found; defaulted employee to Waive."})
        if not tier:
            if employee_election == "Waive":
                tier = "Waive"
            elif raw_plan_name and options.infer_dependents_enrolled_when_tier_missing:
                tier = _infer_tier_from_dependents(dependents)
                notes.append({"Kind": "Warning", "Issue": f"Input row {idxs[0] + 2}: medical plan exists but no tier was found; inferred '{tier}' from listed dependents."})
            elif raw_plan_name:
                tier = "Employee Only"

        employee = base_row()
        employee["First Name"] = _first_nonblank(rows, cm.employee_first)
        employee["Last Name"] = _first_nonblank(rows, cm.employee_last)
        apply_placeholder(employee, "Employee")
        employee["Employee ID"] = _first_nonblank(rows, cm.employee_id)
        employee["Relationship"] = "Employee"
        employee["DOB"] = _first_nonblank(rows, cm.employee_dob)
        employee["Gender"] = _first_nonblank(rows, cm.employee_gender)
        employee["Email"] = _first_nonblank(rows, cm.employee_email)
        employee["Address Line 1"] = _first_nonblank(rows, cm.address1)
        employee["Address Line 2"] = _first_nonblank(rows, cm.address2)
        employee["City"] = _first_nonblank(rows, cm.city)
        employee["State"] = _first_nonblank(rows, cm.state)
        employee["Zip Code"] = _first_nonblank(rows, cm.zip_code)
        employee["County"] = _first_nonblank(rows, cm.county)
        employee["Primary Worksite Zip Code"] = _first_nonblank(rows, cm.worksite_zip)
        employee["Primary Worksite County"] = _first_nonblank(rows, cm.worksite_county)
        employee["ICHRA Class"] = _derive_ichra_class(first_row, cm)
        employee["Health Election"] = employee_election
        employee["Current Health Plan Vendor"] = _first_nonblank(rows, cm.plan_vendor)
        employee["Current Health Plan"] = raw_plan_name
        employee["Current Health Plan Tier"] = tier if tier != "Waive" else ""
        employee["Current Health Plan OOP (single)"] = _first_nonblank(rows, cm.oop_single)
        employee["Current Health Plan OOP (family)"] = _first_nonblank(rows, cm.oop_family)
        employee["Current Health Plan Deductible (single)"] = _first_nonblank(rows, cm.ded_single)
        employee["Current Health Plan Deductible (family)"] = _first_nonblank(rows, cm.ded_family)
        employee["Current Health Plan ER Cost"] = _first_nonblank(rows, cm.er_cost)
        employee["Current Health Plan EE Cost"] = _first_nonblank(rows, cm.ee_cost)
        employee["Annual Salary"] = _first_nonblank(rows, cm.annual_salary)
        employee["Hourly Rate"] = _first_nonblank(rows, cm.hourly_rate)
        employee["Hours Per Week"] = _first_nonblank(rows, cm.hours_per_week)
        employee["Notes"] = _first_nonblank(rows, cm.notes)
        output_rows.append(employee)

        for dep in dependents:
            dep_row = base_row()
            dep_row["First Name"] = dep.get("First Name", "")
            dep_row["Last Name"] = dep.get("Last Name", "")
            dep_row["Relationship"] = dep.get("Relationship", "")

            # Generic dependent "Name" columns are often first-name only. When
            # a dependent has a name but no last name, inherit the employee last
            # name. If the single name cell appears to contain a full name, split
            # the final token into Last Name first.
            if dep_row["First Name"] and not dep_row["Last Name"]:
                name_parts = dep_row["First Name"].split()
                if len(name_parts) >= 2:
                    dep_row["First Name"] = " ".join(name_parts[:-1])
                    dep_row["Last Name"] = name_parts[-1]
                elif employee.get("Last Name"):
                    dep_row["Last Name"] = employee["Last Name"]

            apply_placeholder(dep_row, dep_row["Relationship"] or "Dependent")
            dep_row["DOB"] = dep.get("DOB", "")
            dep_row["Zip Code"] = employee["Zip Code"]
            dep_row["County"] = employee["County"]
            if options.inherit_worksite_zip_to_dependents:
                dep_row["Primary Worksite Zip Code"] = employee["Primary Worksite Zip Code"]
                dep_row["Primary Worksite County"] = employee["Primary Worksite County"]
            dep_row["ICHRA Class"] = employee["ICHRA Class"]
            dep_row["Health Election"] = dependent_election_from_tier(employee_election, tier, dep_row["Relationship"])
            output_rows.append(dep_row)
            if not dep_row["Relationship"]:
                notes.append({"Kind": "Warning", "Issue": f"Input row {dep.get('_source_row', '?')}: dependent has no relationship."})
            if not dep_row["DOB"]:
                notes.append({"Kind": "Warning", "Issue": f"Input row {dep.get('_source_row', '?')}: dependent has no DOB."})

    converted = pd.DataFrame(output_rows, columns=OUTPUT_COLUMNS)
    return {
        "converted_df": converted,
        "horizontal_notes": notes,
        "source_rows": len(df),
        "output_rows": len(converted),
        "employee_groups": len(groups),
        "rowwise_horizontal": rowwise,
    }


# =============================================================================
# 8. VALIDATION ENGINE
# =============================================================================

def validate_date(value: str, *, enforce_required: bool = True) -> dict[str, Any]:
    original = str(value or "").strip()
    if not original:
        if enforce_required:
            return {"ok": False, "msg": "DOB is required", "fixed_val": "", "fix_note": None}
        return {"ok": True, "msg": None, "fixed_val": "", "fix_note": None}

    if re.fullmatch(r"\d{5}", original):
        try:
            parsed = pd.to_datetime(float(original), unit="D", origin="1899-12-30")
            fixed = parsed.strftime("%m/%d/%Y")
            return {"ok": True, "msg": None, "fixed_val": fixed, "fix_note": f"Excel date serial converted: '{original}' -> '{fixed}'"}
        except Exception:
            pass

    parsed = pd.to_datetime(original, errors="coerce")
    if pd.notna(parsed):
        fixed = parsed.strftime("%m/%d/%Y")
        year = int(parsed.year)
        if not 1900 <= year <= 2100:
            return {"ok": False, "msg": f"Year out of range ({year}) in date '{original}'", "fixed_val": fixed, "fix_note": None}
        note = f"Date reformatted: '{original}' -> '{fixed}'" if fixed != original else None
        return {"ok": True, "msg": None, "fixed_val": fixed, "fix_note": note}

    return {"ok": False, "msg": f"Invalid date format '{original}' - expected mm/dd/yyyy", "fixed_val": original, "fix_note": None}


def validate_cell(field_def: dict[str, Any], raw: str, *, enforce_required: bool = True) -> dict[str, Any]:
    value = str(raw or "").strip()
    if field_def["required"] and enforce_required and value == "":
        return {"ok": False, "msg": f"{field_def['name']} is required", "fixed_val": "", "fix_note": None}
    if value == "":
        return {"ok": True, "msg": None, "fixed_val": "", "fix_note": None}

    field_type = field_def["type"]
    if field_type == "date":
        return validate_date(value, enforce_required=enforce_required)
    if field_type == "relationship":
        canonical, note = normalize_relationship(value)
        if canonical not in {"Employee", "Spouse", "Child"}:
            return {"ok": False, "msg": f"Invalid Relationship '{value}' - expected Employee, Spouse, or Child", "fixed_val": canonical, "fix_note": note}
        return {"ok": True, "msg": None, "fixed_val": canonical, "fix_note": note}
    if field_type == "election":
        canonical, note = normalize_election(value)
        if canonical not in {"Enroll", "Waive"}:
            return {"ok": False, "msg": f"Invalid Health Election '{value}' - expected Enroll or Waive", "fixed_val": value, "fix_note": None}
        return {"ok": True, "msg": None, "fixed_val": canonical, "fix_note": note}
    if field_type == "tier":
        canonical, note = normalize_tier(value)
        valid = {"Employee Only", "Employee + Spouse", "Employee + Children", "Family", "Waive", ""}
        if canonical in valid:
            return {"ok": True, "msg": None, "fixed_val": "" if canonical == "Waive" else canonical, "fix_note": note}
        return {"ok": False, "msg": f"Invalid value '{value}' for {field_def['name']} - expected Employee Only, Employee + Spouse, Employee + Children, or Family", "fixed_val": value, "fix_note": None}
    if field_type == "zipcode":
        us_state_tokens = {
            "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA",
            "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
            "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT",
            "VA", "WA", "WV", "WI", "WY", "DC", "PR", "VI", "GU", "AS", "MP",
        }
        if value.upper() in us_state_tokens:
            return {"ok": False, "msg": f"{field_def['name']}: looks like a state abbreviation ('{value}'), not a ZIP code", "fixed_val": value, "fix_note": None}
        fixed = _pad_zip(value)
        note = f"Zip standardized: '{value}' -> '{fixed}'" if fixed != value else None
        if not re.fullmatch(r"\d{5}", fixed):
            if re.search(r"[A-Za-z]", value):
                msg = f"{field_def['name']} appears to be a non-U.S. postal code ('{value}'); expected a U.S. 5-digit ZIP code"
            else:
                msg = f"{field_def['name']} must be a 5-digit numeric ZIP, got '{value}'"
            return {"ok": False, "msg": msg, "fixed_val": fixed, "fix_note": note}
        if fixed in KNOWN_INVALID_OR_UNSUPPORTED_ZIPS:
            return {"ok": False, "msg": f"{field_def['name']} '{fixed}' is invalid", "fixed_val": fixed, "fix_note": note}
        return {"ok": True, "msg": None, "fixed_val": fixed, "fix_note": note}
    if field_type == "numeric":
        clean = value.replace(",", "").replace("$", "").strip()
        if clean.endswith("%"):
            clean = clean[:-1].strip()
        if not re.fullmatch(r"-?\d+(\.\d+)?", clean):
            return {"ok": False, "msg": f"{field_def['name']} must be numeric, got '{value}'", "fixed_val": value, "fix_note": None}
        note = f"Stripped formatting: '{value}' -> '{clean}'" if clean != value else None
        return {"ok": True, "msg": None, "fixed_val": clean, "fix_note": note}
    return {"ok": True, "msg": None, "fixed_val": value, "fix_note": None}


def _raw_for_field(row: pd.Series, col_map: dict[int, int], field_name: str) -> str:
    field_idx = FIELD_IDX_BY_NAME[field_name]
    col_idx = col_map.get(field_idx)
    if col_idx is None:
        return ""
    return str(row.iloc[col_idx] or "").strip()


def _append_issue(target: list[dict[str, Any]], row: Any, col: Any, field: str, value: str, kind: str, issue: str) -> None:
    target.append({"Row": row, "Col": col, "Field": field, "Value": value, "Kind": kind, "Issue": issue})


def _recognized_tier_or_blank(value: object) -> tuple[str, Optional[str]]:
    """Return a valid canonical health-plan tier or blank.

    The quoting import only accepts Employee Only, Employee + Spouse,
    Employee + Children, and Family. Waive is accepted as an input synonym but
    exported as blank because the CSA template uses Health Election for waive.
    """
    tier, note = normalize_tier(value)
    if tier in VALID_HEALTH_PLAN_TIERS:
        return tier, note
    if tier == "Waive":
        return "", note
    return "", None


def _clean_non_election_values_before_family_inference(rows: list[dict[str, str]], fixes: list[dict[str, Any]]) -> None:
    """Treat non-election values in Health Election as missing when a plan exists.

    Some files shift or duplicate the plan name into Health Election. If a
    Current Health Plan is present, those values should not remain as invalid
    elections. Clearing them allows the employee row to default to Enroll and
    dependent rows to derive Enroll/Waive from the employee tier.
    """
    for idx, row in enumerate(rows):
        output_row_num = idx + 2
        raw_election = str(row.get("Health Election", "") or "").strip()
        if not raw_election:
            continue
        canonical_election, election_note = normalize_election(raw_election)
        if canonical_election in {"Enroll", "Waive"}:
            if canonical_election != raw_election:
                row["Health Election"] = canonical_election
                _append_issue(
                    fixes,
                    output_row_num,
                    "Q",
                    "Health Election",
                    raw_election,
                    "Auto-fix",
                    election_note or f"Health Election normalized: '{raw_election}' -> '{canonical_election}'",
                )
            continue
        if str(row.get("Current Health Plan", "") or "").strip():
            row["Health Election"] = ""
            _append_issue(
                fixes,
                output_row_num,
                "Q",
                "Health Election",
                raw_election,
                "Auto-fix",
                f"Health Election value '{raw_election}' is not Enroll/Waive; treated as missing because a Current Health Plan is present.",
            )


def _clean_non_tier_values_before_family_inference(rows: list[dict[str, str]], fixes: list[dict[str, Any]]) -> None:
    """Treat non-tier values in Current Health Plan Tier as missing.

    Some employer files put the medical plan name into both Current Health Plan
    and Current Health Plan Tier. Without this pass, the validator sees a
    non-empty tier and never infers Employee Only / Employee + Spouse /
    Employee + Children / Family from the vertical family unit. This function
    canonicalizes real tier values and clears anything that is not a tier when
    a Current Health Plan is present.
    """
    for idx, row in enumerate(rows):
        output_row_num = idx + 2
        raw_tier = str(row.get("Current Health Plan Tier", "") or "").strip()
        if not raw_tier:
            continue

        canonical_tier, tier_note = _recognized_tier_or_blank(raw_tier)
        if canonical_tier:
            if canonical_tier != raw_tier:
                row["Current Health Plan Tier"] = canonical_tier
                _append_issue(
                    fixes,
                    output_row_num,
                    "T",
                    "Current Health Plan Tier",
                    raw_tier,
                    "Auto-fix",
                    tier_note or f"Current Health Plan Tier normalized: '{raw_tier}' -> '{canonical_tier}'",
                )
            continue

        # If a plan exists but the tier value is not an accepted tier, treat it
        # as a missing tier. This handles duplicated plan names and other plan
        # descriptors that accidentally land in the tier column.
        if str(row.get("Current Health Plan", "") or "").strip():
            row["Current Health Plan Tier"] = ""
            _append_issue(
                fixes,
                output_row_num,
                "T",
                "Current Health Plan Tier",
                raw_tier,
                "Auto-fix",
                f"Current Health Plan Tier value '{raw_tier}' is not a recognized tier; treated as missing so the family-unit tier can be inferred.",
            )


def _employee_level_health_fields() -> list[str]:
    return [
        "Current Health Plan Vendor",
        "Current Health Plan",
        "Current Health Plan Tier",
        "Current Health Plan OOP (single)",
        "Current Health Plan OOP (family)",
        "Current Health Plan Deductible (single)",
        "Current Health Plan Deductible (family)",
        "Current Health Plan ER Cost",
        "Current Health Plan EE Cost",
    ]


def _clear_dependent_employee_level_plan_fields(rows: list[dict[str, str]], fixes: list[dict[str, Any]]) -> None:
    """Clear employee-level current-plan fields on spouse/child rows.

    The CSA import template expects current-plan/cost comparison details at the
    employee row level. Dependents keep Relationship, DOB, ZIP/worksite fields,
    ICHRA Class, and Health Election, while plan/tier/cost fields remain blank.
    """
    fields_to_clear = _employee_level_health_fields()
    for idx, row in enumerate(rows):
        if row.get("Relationship") not in {"Spouse", "Child"}:
            continue
        output_row_num = idx + 2
        for field_name in fields_to_clear:
            existing = str(row.get(field_name, "") or "").strip()
            if existing:
                row[field_name] = ""
                _append_issue(
                    fixes,
                    output_row_num,
                    "—",
                    field_name,
                    existing,
                    "Auto-fix",
                    f"Cleared employee-level {field_name} from dependent row.",
                )


def _vertical_family_units(rows: list[dict[str, str]]) -> list[tuple[int, list[int]]]:
    """Return employee rows paired with their following dependent rows.

    A vertical CSA census represents a family unit as one Employee row followed by
    zero or more Spouse/Child rows. This helper keeps that grouping explicit so
    tier inference can look at the entire family unit before dependent elections
    are derived.
    """
    units: list[tuple[int, list[int]]] = []
    current_employee_idx: Optional[int] = None
    current_dependent_idxs: list[int] = []

    for idx, row in enumerate(rows):
        relationship = row.get("Relationship", "")
        if relationship == "Employee":
            if current_employee_idx is not None:
                units.append((current_employee_idx, current_dependent_idxs))
            current_employee_idx = idx
            current_dependent_idxs = []
        elif relationship in {"Spouse", "Child"} and current_employee_idx is not None:
            current_dependent_idxs.append(idx)

    if current_employee_idx is not None:
        units.append((current_employee_idx, current_dependent_idxs))

    return units


def _describe_dependent_mix(dependents: list[dict[str, str]]) -> str:
    spouse_count = sum(1 for dep in dependents if dep.get("Relationship") == "Spouse")
    child_count = sum(1 for dep in dependents if dep.get("Relationship") == "Child")
    parts: list[str] = []
    if spouse_count:
        parts.append(f"{spouse_count} spouse" if spouse_count == 1 else f"{spouse_count} spouses")
    if child_count:
        parts.append(f"{child_count} child" if child_count == 1 else f"{child_count} children")
    return " and ".join(parts) if parts else "no spouse/child dependents"


def _infer_missing_vertical_employee_tiers(rows: list[dict[str, str]], fixes: list[dict[str, Any]]) -> None:
    """Infer missing employee Current Health Plan Tier from vertical family units.

    When an employee has Current Health Plan populated but Current Health Plan
    Tier is blank, the safest available assumption is the coverage composition
    represented by that employee's family unit:
    - Employee only: no spouse/child rows
    - Employee + Spouse: spouse row and no child rows
    - Employee + Children: child row(s) and no spouse row
    - Family: spouse row and child row(s)
    """
    for employee_idx, dependent_idxs in _vertical_family_units(rows):
        employee = rows[employee_idx]
        output_row_num = employee_idx + 2
        if employee.get("Current Health Plan Tier", "").strip():
            continue
        if not employee.get("Current Health Plan", "").strip():
            continue
        if employee.get("Health Election", "") == "Waive":
            continue

        dependents = [
            {"Relationship": rows[dep_idx].get("Relationship", "")}
            for dep_idx in dependent_idxs
            if rows[dep_idx].get("Relationship", "") in {"Spouse", "Child"}
        ]
        inferred_tier = _infer_tier_from_dependents(dependents)
        employee["Current Health Plan Tier"] = inferred_tier
        mix_description = _describe_dependent_mix(dependents)
        _append_issue(
            fixes,
            output_row_num,
            "T",
            "Current Health Plan Tier",
            "",
            "Auto-fix",
            f"Employee Current Health Plan Tier inferred as '{inferred_tier}' because Current Health Plan is present, tier is missing, and the family unit includes {mix_description}.",
        )


def _apply_vertical_inheritance(rows: list[dict[str, str]], fixes: list[dict[str, Any]], warnings: list[dict[str, Any]]) -> None:
    current_employee: Optional[dict[str, str]] = None
    current_employee_row_num: Optional[int] = None
    inherited_fields = ["Zip Code", "County", "Primary Worksite Zip Code", "Primary Worksite County", "ICHRA Class"]

    # First clear/normalize values that are not real elections or tiers, such as
    # files where the plan name was duplicated into Health Election and/or
    # Current Health Plan Tier. Then infer the employee tier before walking
    # dependents so dependent Health Election can be derived from the completed
    # Employee Only / Employee + Spouse / Employee + Children / Family tier.
    _clean_non_election_values_before_family_inference(rows, fixes)
    _clean_non_tier_values_before_family_inference(rows, fixes)
    _infer_missing_vertical_employee_tiers(rows, fixes)

    for position, row in enumerate(rows):
        output_row_num = position + 2
        relationship = row.get("Relationship", "")
        if relationship == "Employee":
            if not row.get("Health Election"):
                if row.get("Current Health Plan") or row.get("Current Health Plan Tier"):
                    row["Health Election"] = "Enroll"
                    _append_issue(fixes, output_row_num, "Q", "Health Election", "", "Auto-fix", "Employee Health Election defaulted to Enroll because a current health plan/tier is present")
                else:
                    row["Health Election"] = "Waive"
                    _append_issue(fixes, output_row_num, "Q", "Health Election", "", "Auto-fix", "Employee Health Election defaulted to Waive because no current health plan/tier is present")
            current_employee = row
            current_employee_row_num = output_row_num
            continue

        if relationship in {"Spouse", "Child"}:
            if not current_employee:
                _append_issue(warnings, output_row_num, "—", "Dependent inheritance", "", "Warning", "Dependent row appears before any employee row, so employee-level fields could not be inherited")
                continue
            for field_name in inherited_fields:
                if not row.get(field_name) and current_employee.get(field_name):
                    row[field_name] = current_employee[field_name]
                    _append_issue(fixes, output_row_num, "—", field_name, "", "Auto-fix", f"Inherited {field_name} from employee row {current_employee_row_num}")
            if not row.get("Health Election"):
                employee_election = current_employee.get("Health Election", "")
                tier = current_employee.get("Current Health Plan Tier", "")
                derived = dependent_election_from_tier(employee_election, tier, relationship)
                row["Health Election"] = derived
                _append_issue(fixes, output_row_num, "Q", "Health Election", "", "Auto-fix", f"Dependent Health Election derived from employee tier '{tier or 'not provided'}' as '{derived}'")

    _clear_dependent_employee_level_plan_fields(rows, fixes)


def run_validation(df: pd.DataFrame, zip_cache: dict[str, dict[str, str]], *, source_label: str = "input") -> dict[str, Any]:
    df = df.fillna("").astype(str).reset_index(drop=True)
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    fixes: list[dict[str, Any]] = []
    clears: list[dict[str, Any]] = []

    col_map, unrecognized_cols, mapping_notes = map_columns(df)
    for column_name in unrecognized_cols:
        # Suppress obvious ancillary columns to keep reports useful, but keep non-ancillary unmapped fields visible.
        if _is_ancillary_header(column_name) or _norm(column_name) in {"date of hire", "active or cobra", "location division", "job title"}:
            continue
        _append_issue(warnings, "—", "—", column_name, "", "Warning", f"Unrecognized column '{column_name}' - not mapped to any canonical field")

    out_rows: list[dict[str, str]] = []
    source_row_numbers: list[int] = []

    for df_index, row in df.iterrows():
        if all(_is_blank(v) for v in row):
            continue
        display_row = int(df_index) + 2
        new_row = {column: "" for column in OUTPUT_COLUMNS}
        raw_by_field: dict[str, str] = {}

        for field_idx, field_def in enumerate(FIELDS):
            field_name = field_def["name"]
            col_idx = col_map.get(field_idx)
            raw = str(row.iloc[col_idx]).strip() if col_idx is not None else ""
            raw_by_field[field_name] = raw

            if field_name in CLEARED_FIELDS:
                if raw:
                    _append_issue(clears, display_row, field_def["col"], field_name, raw, "Cleared", f"Field '{field_name}' automatically cleared")
                new_row[field_name] = ""
                continue

            result = validate_cell(field_def, raw, enforce_required=False)
            if result["fix_note"]:
                _append_issue(fixes, display_row, field_def["col"], field_name, raw, "Auto-fix", result["fix_note"])
            new_row[field_name] = result["fixed_val"] if result["ok"] else raw

        # Medical Election values like EE/ES/EC/EF often represent both enrollment and tier.
        if not new_row["Current Health Plan Tier"]:
            tier, tier_note = derive_tier_from_coverage_code(raw_by_field.get("Health Election", ""))
            if tier:
                new_row["Current Health Plan Tier"] = tier
                _append_issue(fixes, display_row, "T", "Current Health Plan Tier", raw_by_field.get("Health Election", ""), "Auto-fix", tier_note or f"Current Health Plan Tier derived as '{tier}'")

        relationship = new_row.get("Relationship", "")
        if relationship == "Employee" and not new_row.get("Health Election"):
            if new_row.get("Current Health Plan") or new_row.get("Current Health Plan Tier"):
                new_row["Health Election"] = "Enroll"
                _append_issue(fixes, display_row, "Q", "Health Election", "", "Auto-fix", "Employee Health Election defaulted to Enroll because a current health plan/tier is present")
            else:
                new_row["Health Election"] = "Waive"
                _append_issue(fixes, display_row, "Q", "Health Election", "", "Auto-fix", "Employee Health Election defaulted to Waive because no current health plan/tier is present")

        out_rows.append(new_row)
        source_row_numbers.append(display_row)

    _apply_vertical_inheritance(out_rows, fixes, warnings)

    # Final validation after normalization and inheritance.
    error_positions: list[int] = []
    valid_positions: list[int] = []
    rows_with_error: set[int] = set()
    zip_fields_with_errors: dict[int, set[str]] = {}

    for position, row in enumerate(out_rows):
        display_row = position + 2
        for field_def in FIELDS:
            field_name = field_def["name"]
            value = row.get(field_name, "")
            result = validate_cell(field_def, value, enforce_required=True)
            if not result["ok"]:
                _append_issue(errors, display_row, field_def["col"], field_name, value, "Error", result["msg"])
                rows_with_error.add(position)
                if field_def["type"] == "zipcode":
                    zip_fields_with_errors.setdefault(position, set()).add(field_name)
            else:
                row[field_name] = result["fixed_val"]

    if zip_cache:
        for position, row in enumerate(out_rows):
            display_row = position + 2
            for zip_field, county_field, col_letter in [
                ("Zip Code", "County", "L"),
                ("Primary Worksite Zip Code", "Primary Worksite County", "N"),
            ]:
                zip_value = row.get(zip_field, "").strip()
                if not zip_value or zip_field in zip_fields_with_errors.get(position, set()):
                    continue
                if not re.fullmatch(r"\d{5}", zip_value):
                    continue
                info = lookup_zip(zip_value, zip_cache)
                if info:
                    if not row.get(county_field, "").strip() and info.get("county"):
                        row[county_field] = info["county"]
                        _append_issue(fixes, display_row, "—", county_field, "", "Auto-fix", f"County auto-filled from ZIP '{zip_value}': '{info['county']}'")
                else:
                    _append_issue(errors, display_row, col_letter, zip_field, zip_value, "Error", f"{zip_field} '{zip_value}' is not a recognized U.S. ZIP code")
                    rows_with_error.add(position)

    for position in range(len(out_rows)):
        if position in rows_with_error:
            error_positions.append(position)
        else:
            valid_positions.append(position)

    for position, row in enumerate(out_rows):
        display_row = position + 2
        er_value = row.get("Current Health Plan ER Cost", "").strip()
        ee_value = row.get("Current Health Plan EE Cost", "").strip()
        if bool(er_value) ^ bool(ee_value):
            _append_issue(warnings, display_row, "Y/Z", "Cost Comparison", "", "Warning", "Both ER Cost and EE Cost should be present for Cost Comparison")

    reformatted_df = pd.DataFrame(out_rows, columns=OUTPUT_COLUMNS).reset_index(drop=True)
    return {
        "errors": errors,
        "warnings": warnings,
        "fixes": fixes,
        "clears": clears,
        "reformatted_df": reformatted_df,
        "valid_positions": valid_positions,
        "error_positions": error_positions,
        "total_rows": len(out_rows),
        "mapping_notes": mapping_notes,
        "unrecognized_cols": unrecognized_cols,
        "source_label": source_label,
    }


# =============================================================================
# 9. DOWNLOAD HELPERS AND SAMPLE DATA
# =============================================================================

def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def issues_to_csv_bytes(errors, warnings, fixes, clears, horizontal_notes=None) -> bytes:
    cols = ["Row", "Col", "Field", "Value", "Kind", "Issue"]
    combined = list(errors) + list(warnings) + list(fixes) + list(clears)
    if horizontal_notes:
        for note in horizontal_notes:
            combined.append({
                "Row": "—",
                "Col": "—",
                "Field": "Horizontal conversion",
                "Value": "",
                "Kind": note.get("Kind", "Note"),
                "Issue": note.get("Issue", ""),
            })
    if not combined:
        return pd.DataFrame(columns=cols).to_csv(index=False).encode("utf-8")
    combined.sort(key=lambda item: (str(item.get("Row", "0")).zfill(6), item.get("Kind", ""), item.get("Field", "")))
    return pd.DataFrame(combined, columns=cols).to_csv(index=False).encode("utf-8")


SAMPLE_VERTICAL_CSV = """First Name,Last Name,"Relationship Employee (EE), Spouse (SP), or Child (CH)",Date of Birth,Home Zip Code,Work Zip Code,Employee Class (Name or #),Weekly Hours Worked,Medical Election,Medical Plan Name,Current Health Plan Tier,Annual Salary
Musa T.,Abdelhadi,Employee,06/11/1980,75126,75243,Full Time Benefits Eligible,40,Enroll,Buy-Up PPO,,127000.02
Neveen,Shalabi,Spouse,08/01/1979,,,,,,,,
Mohammod,Abdelhadi,Child,11/16/2002,,,,,,,,
Danya,Abdelhadi,Child,04/12/2004,,,,,,,,
Alex,EmployeeOnly,Employee,02/10/1984,53201,53201,Full Time Benefits Eligible,40,Enroll,Base PPO,,82000
Jamie,SpouseOnly,Employee,03/12/1979,60601,60601,Full Time Benefits Eligible,40,Enroll,Base PPO,,91000
Jordan,SpouseOnly,Spouse,08/15/1980,,,,,,,,
Taylor,ChildrenOnly,Employee,04/14/1986,75243,75243,Full Time Benefits Eligible,40,Enroll,Base PPO,,87000
Morgan,ChildrenOnly,Child,09/22/2014,,,,,,,,
Riley,ChildrenOnly,Child,01/05/2017,,,,,,,,
"""

SAMPLE_HORIZONTAL_CSV = """BIRTH DATE,PRIMARY ADDRESS LINE 1,PRIMARY ADDRESS - CITY,PRIMARY ADDRESS - STATE / TERRITORY,PRIMARY ADDRESS - ZIP CODE,WORKSITE ZIP CODE,WORKER CATEGORY,ANNUAL SALARY,MEDICAL PLAN NAME,MEDICAL COVERAGE TIER,Dependent Number,Dependent Relationship,Dependent DOB
11/26/1986,10502 Barrichello St,Bakersfield,CA,93314,93314,F - Full Time,89139.96,HDHP HSA PRIME PLUS,Family,1,Child,3/4/2011
11/26/1986,10502 Barrichello St,Bakersfield,CA,93314,93314,F - Full Time,89139.96,HDHP HSA PRIME PLUS,Family,2,Child,11/2/2008
11/26/1986,10502 Barrichello St,Bakersfield,CA,93314,93314,F - Full Time,89139.96,HDHP HSA PRIME PLUS,Family,3,Spouse,6/30/1978
1/26/1989,10503 Barrichello St,Bakersfield,CA,93314,93314,F - Full Time,133240.64,HDHP HSA PRIME PLUS,Employee + Children,1,Child,2/8/2014
1/26/1989,10503 Barrichello St,Bakersfield,CA,93314,93314,F - Full Time,133240.64,HDHP HSA PRIME PLUS,Employee + Children,2,Child,2/16/2008
"""


# =============================================================================
# 10. STREAMLIT UI
# =============================================================================

def _style_issues(issue_rows: list[dict[str, Any]]) -> "pd.io.formats.style.Styler":
    cols = ["Row", "Col", "Field", "Value", "Kind", "Issue"]
    df = pd.DataFrame(issue_rows, columns=cols)
    if df.empty:
        df = pd.DataFrame(columns=cols)
    df = df.sort_values(["Row", "Kind"], key=lambda series: series.map(lambda x: str(x).zfill(6) if str(x).isdigit() else str(x))).reset_index(drop=True)

    def color_kind(value: str) -> str:
        if value == "Error":
            return "color:#993C1D;font-weight:600"
        if value == "Warning":
            return "color:#854F0B;font-weight:600"
        if value == "Auto-fix":
            return "color:#3B6D11;font-weight:600"
        if value == "Cleared":
            return "color:#555599;font-weight:600"
        return ""

    styler = df.style
    return styler.map(color_kind, subset=["Kind"]) if hasattr(styler, "map") else styler.applymap(color_kind, subset=["Kind"])


def _build_results(df_input: pd.DataFrame, input_mode: str, use_zip_lookup: bool, horizontal_options: HorizontalOptions) -> dict[str, Any]:
    zip_cache = build_zip_cache() if (use_zip_lookup and ZIPCODES_AVAILABLE) else {}
    horizontal_detected, detect_reason = detect_horizontal_census(df_input)
    force_horizontal = input_mode == "Horizontal dependent census"
    force_vertical = input_mode == "Standard vertical census"
    use_horizontal = force_horizontal or (horizontal_detected and not force_vertical)

    horizontal_notes: list[dict[str, str]] = []
    conversion_summary: Optional[dict[str, Any]] = None
    validation_input = df_input
    source_label = "standard vertical census"

    if use_horizontal:
        conversion = convert_horizontal_to_vertical(df_input, horizontal_options)
        validation_input = conversion["converted_df"]
        horizontal_notes = conversion["horizontal_notes"]
        conversion_summary = conversion
        source_label = "horizontal dependent census"

    results = run_validation(validation_input, zip_cache, source_label=source_label)
    reformatted_df = results["reformatted_df"]
    results["valid_df"] = reformatted_df.iloc[results["valid_positions"]].reset_index(drop=True)
    results["error_df"] = reformatted_df.iloc[results["error_positions"]].reset_index(drop=True)
    results["horizontal_used"] = use_horizontal
    results["detect_reason"] = detect_reason
    results["horizontal_notes"] = horizontal_notes
    results["conversion_summary"] = conversion_summary
    return results


def main() -> None:
    st.set_page_config(page_title="Census Validator", layout="wide")
    st.markdown(
        """
        <style>
            .block-container { padding-top: 1.4rem; max-width: 1240px; }
            [data-testid="metric-container"] { background:#f7f7f5; border-radius:8px; padding:10px 14px; }
            .banner-ok { background:#EAF3DE; border:1px solid #97C459; border-radius:8px; padding:12px 16px; color:#3B6D11; font-size:14px; margin-bottom:1rem; }
            .info-box { background:#EEF2FB; border:1px solid #ADC0EF; border-radius:8px; padding:12px 16px; font-size:13px; color:#1a3a7a; margin-bottom:1rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("Census Validator & Reformatter")
    st.caption("Vertical CSA census validation, horizontal dependent conversion, header/footer row detection, and system-ready CSV export")

    with st.sidebar:
        st.header("Options")
        input_mode = st.radio(
            "Input format",
            ["Auto-detect", "Standard vertical census", "Horizontal dependent census"],
            index=0,
            help="Auto-detect now treats verbose vertical relationship headers correctly.",
        )
        use_zip_lookup = st.toggle("Auto-fill County from ZIP", value=ZIPCODES_AVAILABLE, disabled=not ZIPCODES_AVAILABLE)
        if not ZIPCODES_AVAILABLE:
            st.caption("Install the zipcodes package to enable county auto-fill.")

        st.divider()
        st.markdown("**Horizontal conversion**")
        auto_placeholders = st.toggle("Create placeholder names when missing", value=True)
        infer_missing_tier = st.toggle("Infer dependent enrollment when medical tier is missing", value=True)
        inherit_worksite = st.toggle("Copy worksite ZIP to dependent rows", value=True)
        exclude_interns = st.toggle("Skip interns/contractors in horizontal mode", value=True)

        st.divider()
        st.markdown("**Required output fields**")
        for field_def in FIELDS:
            marker = " required" if field_def["required"] else ""
            cleared = " cleared" if field_def["name"] in CLEARED_FIELDS else ""
            st.caption(f"{field_def['col']}. {field_def['name']}{marker}{cleared}")

    col_upload, col_paste = st.columns([1, 1], gap="large")
    with col_upload:
        uploaded = st.file_uploader("Upload census file (.csv / .xlsx / .xls)", type=["csv", "xlsx", "xls"])
    with col_paste:
        pasted = st.text_area("or paste CSV data", value=st.session_state.get("paste_text", ""), height=170, placeholder="Paste CSV including the header row")

    button_col1, button_col2, button_col3, _ = st.columns([1, 1, 1, 5])
    with button_col1:
        run_btn = st.button("Validate", type="primary", use_container_width=True)
    with button_col2:
        if st.button("Load vertical sample", use_container_width=True):
            st.session_state["paste_text"] = SAMPLE_VERTICAL_CSV
            st.rerun()
    with button_col3:
        if st.button("Load horizontal sample", use_container_width=True):
            st.session_state["paste_text"] = SAMPLE_HORIZONTAL_CSV
            st.rerun()

    if run_btn:
        df_input: Optional[pd.DataFrame] = None
        try:
            if uploaded:
                df_input = read_uploaded_file(uploaded)
            elif pasted.strip():
                df_input = _read_csv_text(pasted)
            else:
                st.warning("Please upload a file or paste CSV data first.")
        except Exception as exc:
            st.error(f"Could not read input: {exc}")

        if df_input is not None:
            with st.spinner("Validating and reformatting..."):
                options = HorizontalOptions(
                    auto_placeholder_names=auto_placeholders,
                    infer_dependents_enrolled_when_tier_missing=infer_missing_tier,
                    inherit_worksite_zip_to_dependents=inherit_worksite,
                    exclude_interns_and_contractors=exclude_interns,
                )
                st.session_state["results"] = _build_results(df_input, input_mode, use_zip_lookup, options)

    if "results" not in st.session_state:
        return

    results = st.session_state["results"]
    errors = results["errors"]
    warnings = results["warnings"]
    fixes = results["fixes"]
    clears = results["clears"]
    horizontal_notes = results.get("horizontal_notes", [])
    reformatted_df = results["reformatted_df"]
    valid_df = results["valid_df"]
    error_df = results["error_df"]
    total = results["total_rows"]
    error_rows = len(results["error_positions"])
    valid_rows = total - error_rows

    st.divider()
    if results.get("horizontal_used"):
        conversion = results.get("conversion_summary") or {}
        st.markdown(
            f'<div class="info-box">Horizontal conversion applied: '
            f'{conversion.get("source_rows", "?")} source rows -> {conversion.get("employee_groups", "?")} employee groups -> '
            f'{conversion.get("output_rows", total)} output rows. Dependents were ordered Employee, Spouse, Child.</div>',
            unsafe_allow_html=True,
        )
    else:
        st.caption(f"Input handled as standard vertical census. Detection note: {results.get('detect_reason', '')}")

    if results.get("mapping_notes") or results.get("unrecognized_cols"):
        with st.expander("Column mapping notes", expanded=False):
            for note in results.get("mapping_notes", []):
                st.caption(f"- {note}")
            visible_unrecognized = [c for c in results.get("unrecognized_cols", []) if not _is_ancillary_header(c)]
            for column in visible_unrecognized:
                st.caption(f"- Unrecognized: `{column}`")

    metric_cols = st.columns(8)
    metric_cols[0].metric("Total rows", total)
    metric_cols[1].metric("Valid rows", valid_rows)
    metric_cols[2].metric("Error rows", error_rows)
    metric_cols[3].metric("Errors", len(errors))
    metric_cols[4].metric("Warnings", len(warnings) + len([n for n in horizontal_notes if n.get("Kind") == "Warning"]))
    metric_cols[5].metric("Auto-fixes", len(fixes) + len([n for n in horizontal_notes if n.get("Kind") == "Auto-fix"]))
    metric_cols[6].metric("Cleared fields", len(clears))
    metric_cols[7].metric("Mode", "Horizontal" if results.get("horizontal_used") else "Vertical")

    st.divider()
    st.subheader("Downloads")
    download_cols = st.columns(4)
    with download_cols[0]:
        st.markdown("##### Full reformatted census")
        st.caption(f"All {total} rows")
        st.download_button("Download full CSV", data=df_to_csv_bytes(reformatted_df), file_name="census_reformatted.csv", mime="text/csv", use_container_width=True, key="dl_full")
    with download_cols[1]:
        st.markdown("##### Valid rows only")
        st.caption(f"{valid_rows} rows")
        st.download_button("Download valid CSV", data=df_to_csv_bytes(valid_df), file_name="census_valid_rows.csv", mime="text/csv", use_container_width=True, key="dl_valid")
    with download_cols[2]:
        st.markdown("##### Error rows only")
        st.caption(f"{error_rows} rows")
        st.download_button("Download error CSV", data=df_to_csv_bytes(error_df), file_name="census_error_rows.csv", mime="text/csv", use_container_width=True, key="dl_errors")
    with download_cols[3]:
        st.markdown("##### Validation report")
        st.caption("Errors, warnings, fixes, clears")
        st.download_button("Download report", data=issues_to_csv_bytes(errors, warnings, fixes, clears, horizontal_notes), file_name="census_validation_report.csv", mime="text/csv", use_container_width=True, key="dl_report")

    st.divider()
    horizontal_issue_rows = [
        {"Row": "—", "Col": "—", "Field": "Horizontal conversion", "Value": "", "Kind": note.get("Kind", "Note"), "Issue": note.get("Issue", "")}
        for note in horizontal_notes
    ]
    all_issues = errors + warnings + fixes + clears + horizontal_issue_rows
    if not all_issues:
        st.markdown(f'<div class="banner-ok">All {total} rows passed validation with no issues.</div>', unsafe_allow_html=True)
    else:
        tabs = st.tabs([
            f"All ({len(all_issues)})",
            f"Errors ({len(errors)})",
            f"Warnings ({len(warnings) + len([n for n in horizontal_notes if n.get('Kind') == 'Warning'])})",
            f"Auto-fixes ({len(fixes) + len([n for n in horizontal_notes if n.get('Kind') == 'Auto-fix'])})",
            f"Cleared ({len(clears)})",
        ])
        with tabs[0]:
            st.dataframe(_style_issues(all_issues), use_container_width=True, hide_index=True)
        with tabs[1]:
            st.dataframe(_style_issues(errors), use_container_width=True, hide_index=True) if errors else st.success("No errors found.")
        with tabs[2]:
            warning_rows = warnings + [row for row in horizontal_issue_rows if row["Kind"] == "Warning"]
            st.dataframe(_style_issues(warning_rows), use_container_width=True, hide_index=True) if warning_rows else st.success("No warnings found.")
        with tabs[3]:
            fix_rows = fixes + [row for row in horizontal_issue_rows if row["Kind"] == "Auto-fix"]
            st.dataframe(_style_issues(fix_rows), use_container_width=True, hide_index=True) if fix_rows else st.info("No auto-fixes applied.")
        with tabs[4]:
            st.dataframe(_style_issues(clears), use_container_width=True, hide_index=True) if clears else st.info("No fields were cleared.")

    st.divider()
    with st.expander("Preview reformatted output", expanded=False):
        st.dataframe(reformatted_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
