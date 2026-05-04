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

# Initialize taxonomy dictionary
TAXONOMY_OBJECTIONS = dict()

# Standalone Keys
K1 = "OB-TX-001: Relevance / Outside Scope"
K2 = "OB-TX-002: Overbroad as to Time"
K3 = "OB-TX-003: Overbroad as to Subject Matter"
K4 = "OB-TX-004: Vague / Ambiguous / Undefined Terms"
K5 = "OB-TX-005: Lack of Reasonable Particularity"
K6 = "OB-TX-006: Improper Fishing Expedition"
K7 = "OB-TX-007: Undue Burden / Expense"
K8 = "OB-TX-008: More Convenient / Less Burdensome Source"
K9 = "OB-TX-009: Duplicative / Previously Produced"
K10 = "OB-TX-010: Premature / Discovery Ongoing"
K11 = "OB-TX-011: Mandatory Initial Disclosures (TRCP 194)"
K12 = "OB-TX-012: TRCP 193.3 Withholding Statement"
K13 = "OB-TX-013: Attorney-Client Privilege"
K14 = "OB-TX-014: Work Product Privilege"
K15 = "OB-TX-015: Consulting Expert Protection"
K16 = "OB-TX-016: Spousal Privilege"
K17 = "OB-TX-017: Mental Health Records Not at Issue"
K18 = "OB-TX-018: Privacy / Sensitive Personal Info"
K19 = "OB-TX-019: Narrative / Marshaling Proof"
K20 = "OB-TX-020: Marshal All Evidence"
K21 = "OB-TX-021: Medical Opinion from Lay Party"
K22 = "OB-TX-022: Improper Expert Discovery by Interrogatory"
K23 = "OB-TX-023: Beyond Current Knowledge"
K24 = "OB-TX-024: Exceeds Numerical Discovery Plan Limits"
K25 = "OB-TX-025: Not in Possession, Custody, or Control"
K26 = "OB-TX-026: Request Requires Creation of a Document"
K27 = "OB-TX-027: Blank Authorization / Lack of Specificity"
K28 = "OB-TX-028: Medical Authorization Improper / Records Forthcoming"
K29 = "OB-TX-029: Tax Returns / Heightened Financial Privacy"
K30 = "OB-TX-030: Social Security / Identifier Forms"
K31 = "OB-TX-031: Employment Records Limited to Wage Period"
K32 = "OB-TX-032: Phone Records Lack Nexus"
K33 = "OB-TX-033: Collateral Source Rule"
K34 = "OB-TX-034: Medical Billing Creation Expansion"
K35 = "OB-TX-035: Premature Damage Computation"
K36 = "OB-TX-036: Core Issue / Disputed Merits RFA"
K37 = "OB-TX-037: Compound RFA"
K38 = "OB-TX-038: Calls for Legal Conclusion"
K39 = "OB-TX-039: Cannot Admit/Deny After Reasonable Inquiry"
K40 = "OB-TX-040: Prejudicial Terms / Controverted Text"
K41 = "OB-TX-041: All Written Complaints (Overbroad)"
K42 = "OB-TX-042: Ex Parte Provider Communications"

# Single-line strings using escaped quotes to guarantee no parsing syntax errors
V1 = "Plaintiff objects to this request under TRCP 193.2 to the extent it seeks information not relevant to any party's claim or defense and is not within the permissible scope of discovery."
V2 = "Plaintiff objects that this request is overly broad because it is not reasonably limited in time to the matters at issue in this litigation."
V3 = "Plaintiff objects that this request is overly broad because it is not reasonably tailored to include only matters relevant to the issues in dispute."
V4 = "Plaintiff objects that this request is vague and ambiguous because it fails to define the key terms with reasonable certainty, such that Plaintiff cannot determine the exact information sought."
V5 = "Plaintiff objects that this request fails to describe the items or documents sought with reasonable particularity and therefore imposes an improper burden on the responding party."
V6 = "Plaintiff objects that this request constitutes an improper fishing expedition and is not reasonably tailored to obtain information directly relevant to the claims or defenses at issue."
V7 = "Plaintiff objects that complying with this request would impose an undue burden and expense that is completely disproportionate to the likely benefit of the discovery."
V8 = "Plaintiff objects to this request to the extent the information or responsive documents sought are obtainable from a source that is more convenient, less burdensome, or less expensive."
V9 = "Plaintiff objects that this request is unreasonably cumulative or duplicative and, to the extent responsive material exists, it has already been produced or identified in prior productions."
V10 = "Plaintiff objects that this request is premature to the extent it seeks a complete factual or evidentiary statement before discovery is sufficiently developed."
V11 = "Plaintiff objects that this request seeks information through an improper discovery vehicle as the subject matter is directly governed by required initial disclosures under TRCP 194.1."
V12 = "Responsive material has been withheld pursuant to Tex. R. Civ. P. 193.3. The withheld material is responsive to this request and is withheld on the basis of applicable privileges."
V13 = "Plaintiff withholds responsive material protected by the attorney-client privilege and provides this withholding statement pursuant to Rule 193.3."
V14 = "Plaintiff objects and withholds responsive material to the extent it constitutes work product or material prepared in anticipation of litigation under TRCP 192.5."
V15 = "Plaintiff objects to this request to the extent it seeks the identity, mental impressions, or opinions of consulting experts not expected to testify."
V16 = "Plaintiff objects to the extent this request seeks confidential communications between spouses protected by the spousal communications privilege."
V17 = "Plaintiff objects to this request to the extent it seeks highly sensitive mental-health information where Plaintiff has not affirmatively placed their mental condition at issue in this litigation."
V18 = "Plaintiff objects to this request to the extent it seeks highly sensitive personal identifiers or private information without a demonstrated need that is proportional to the case."
V19 = "Plaintiff objects to this interrogatory to the extent it requires a narrative response or detailed marshaling of proof more appropriately developed through deposition."
V20 = "Plaintiff objects that this interrogatory improperly seeks to force Plaintiff to marshal all evidence supporting its claims or defenses."
V21 = "Plaintiff objects to the extent this interrogatory requires a lay party Plaintiff to provide medical opinions beyond Plaintiff's personal knowledge or expert qualifications."
V22 = "Plaintiff objects to this interrogatory to the extent it seeks expert information outside the scope or manner authorized by the expert discovery rules."
V23 = "Plaintiff objects to the extent this interrogatory seeks information beyond Plaintiff's present knowledge and attempts to bind Plaintiff to a complete evidentiary statement before discovery is complete."
V24 = "Plaintiff objects because this set exceeds the maximum number of permissible requests or answers under the governing TRCP discovery control plan."
V25 = "Plaintiff objects to the extent this request seeks materials not within Plaintiff's possession, custody, or control."
V26 = "Plaintiff objects because this request improperly requires Plaintiff to create a document that does not presently exist."
V27 = "Plaintiff objects to signing the requested authorization in blank because it fails to specify the records sought or specific providers, depriving Plaintiff of a meaningful opportunity to evaluate relevance."
V28 = "Plaintiff objects to the extent this request seeks a blanket medical authorization rather than relevant medical records properly subject to production or disclosure."
V29 = "Plaintiff objects that this request seeks highly confidential tax information and is overbroad, intrusive, and not shown to be necessary in the form requested."
V30 = "Plaintiff objects to this request to the extent it seeks Social Security records or identifiers that are irrelevant, overbroad, and unduly intrusive."
V31 = "Plaintiff objects to this request to the extent it seeks personnel and employment records beyond those reasonably related to any claimed wage loss."
V32 = "Plaintiff objects because this request seeks telephone records without a pleaded or otherwise demonstrated nexus to any claim or defense in the case."
V33 = "Plaintiff objects to this request to the extent it seeks information regarding non-utilized health insurance, private health plans, or collateral benefits barred by the Texas Collateral Source Rule."
V34 = "Plaintiff objects to the extent this request improperly expands the burden of medical billing disclosure by requiring the creation of an itemized analysis beyond the records themselves."
V35 = "Plaintiff objects to the extent this request seeks a premature, exhaustive, or artificially fixed statement of damages before discovery is fully developed."
V36 = "Plaintiff objects because this request for admission improperly seeks to establish a disputed merits issue rather than narrow an uncontroverted fact."
