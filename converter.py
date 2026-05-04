import streamlit as st
import os
import io
import re
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfReader
import google.generativeai as genai
from openai import OpenAI

# 1. Page Configuration
st.set_page_config(page_title="Discovery Drafter & Utility", layout="wide")

st.title("Legal Utility: PDF Converter & Discovery Response Drafter")
st.markdown("---")

# 2. Canonical Texas 42-Point Objection Taxonomy
TAXONOMY_OBJECTIONS = {
    # 1. Core Scope and Form Objections
    "OB-TX-001: Relevance / Outside Scope": (
        "Plaintiff objects to this request under TRCP 193.2 to the extent it seeks information "
        "not relevant to any party's claim or defense and is not within the permissible scope of discovery."
    ),
    "OB-TX-002: Overbroad as to Time": (
        "Plaintiff objects that this request is overly broad because it is not reasonably limited in time "
        "to the matters at issue in this litigation."
    ),
    "OB-TX-003: Overbroad as to Subject Matter": (
        "Plaintiff objects that this request is overly broad because it is not reasonably tailored to include "
        "only matters relevant to the issues in dispute."
    ),
    "OB-TX-004: Vague / Ambiguous / Undefined Terms": (
        "Plaintiff objects that this request is vague and ambiguous because it fails to define the key terms "
        "with reasonable certainty, such that Plaintiff cannot determine the exact information sought."
    ),
    "OB-TX-005: Lack of Reasonable Particularity": (
        "Plaintiff objects that this request fails to describe the items or documents sought with reasonable "
        "particularity and therefore imposes an improper burden on the responding party."
    ),
    "OB-TX-006: Improper Fishing Expedition": (
        "Plaintiff objects that this request constitutes an improper fishing expedition and is not reasonably "
        "tailored to obtain information directly relevant to the claims or defenses at issue."
    ),
    
    # 2. Burden, Proportionality, and Procedural Objections
    "OB-TX-007: Undue Burden / Expense": (
        "Plaintiff objects that complying with this request would impose an undue burden and expense "
        "that is completely disproportionate to the likely benefit of the discovery."
    ),
    "OB-TX-008: More Convenient / Less Burdensome Source": (
        "Plaintiff objects to this request to the extent the information or responsive documents sought are "
        "obtainable from a source that is more convenient, less burdensome, or less expensive."
    ),
    "OB-TX-009: Duplicative / Previously Produced": (
        "Plaintiff objects that this request is unreasonably cumulative or duplicative and, to the extent responsive "
        "material exists, it has already been produced or identified in prior productions."
    ),
    "OB-TX-010: Premature / Discovery Ongoing": (
        "Plaintiff objects that this request is premature to the extent it seeks a complete factual or evidentiary "
        "statement before discovery is sufficiently developed."
    ),
    "OB-TX-011: Mandatory Initial Disclosures (TRCP 194)": (
        "Plaintiff objects that this request seeks information through an improper discovery vehicle as the "
        "subject matter is directly governed by required initial disclosures under TRCP 194.1."
    ),
    
    # 3. Privilege and Protected-Information Objections
    "OB-TX-012: TRCP 193.3 Withholding Statement": (
        "Responsive material has been withheld pursuant to Tex. R. Civ. P. 193.3. The withheld material is responsive "
        "to this request and is withheld on the basis of applicable privileges."
    ),
    "OB-TX-013: Attorney-Client Privilege": (
        "Plaintiff withholds responsive material protected by the attorney-client privilege and provides this "
        "withholding statement pursuant to Rule 193.3."
    ),
    "OB-TX-014: Work Product Privilege": (
        "Plaintiff objects and withholds responsive material to the extent it constitutes work product or material "
        "prepared in anticipation of litigation under TRCP 192.5."
    ),
    "OB-TX-015: Consulting Expert Protection": (
        "Plaintiff objects to this request to the extent it seeks the identity, mental impressions, or opinions of "
        "consulting experts not expected to testify."
    ),
    "OB-TX-016: Spousal Privilege": (
        "Plaintiff objects to the extent this request seeks confidential communications between spouses protected "
        "by the spousal communications privilege."
    ),
    "OB-TX-017: Mental Health Records Not at Issue": (
        "Plaintiff objects to this request to the extent it seeks highly sensitive mental-health information where "
        "Plaintiff has not affirmatively placed their mental condition at issue in this litigation."
    ),
    "OB-TX-018: Privacy / Sensitive Personal Info": (
        "Plaintiff objects to this request to the extent it seeks highly sensitive personal identifiers or private "
        "information without a demonstrated need that is proportional to the case."
    ),
    
    # 4. Interrogatory-Specific Objections
    "OB-TX-019: Narrative / Marshaling Proof": (
        "Plaintiff objects to this interrogatory
