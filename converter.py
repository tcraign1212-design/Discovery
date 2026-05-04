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

# Safe strings using triple quotes to prevent any syntax or line-break errors
O1 = r"""Plaintiff objects to this request under TRCP 193.2 to the extent it seeks information not relevant to any party's claim or defense and is not within the permissible scope of discovery."""
O2 = r"""Plaintiff objects that this request is overly broad because it is not reasonably limited in time to the matters at issue in this litigation."""
O3 = r"""Plaintiff objects that this request is overly broad because it is not reasonably tailored to include only matters relevant to the issues in dispute."""
O4 = r"""Plaintiff objects that this request is vague and ambiguous because it fails to define the key terms with reasonable certainty, such that Plaintiff cannot determine the exact information sought."""
O5 = r"""Plaintiff objects that this request fails to describe the items or documents sought with reasonable particularity and therefore imposes an improper burden on the responding party."""
O6 = r"""Plaintiff objects that this request constitutes an improper fishing expedition and is not reasonably tailored to obtain information directly relevant to the claims or defenses at issue."""
O7 = r"""Plaintiff objects that complying with this request would impose an undue burden and expense that is completely disproportionate to the likely benefit of the discovery."""
O8 = r"""Plaintiff objects to this request to the extent the information or responsive documents sought are obtainable from a source that is more convenient, less burdensome, or less expensive."""
O9 = r"""Plaintiff objects that this request is unreasonably cumulative or duplicative and, to the extent responsive material exists, it has already been produced or identified in prior productions."""
O10 = r"""Plaintiff objects that this request is premature to the extent it seeks a complete factual or evidentiary statement before discovery is sufficiently developed."""
O11 = r"""Plaintiff objects that this request seeks information through an improper discovery vehicle as the subject matter is directly governed by required initial disclosures under TRCP 194.1."""
O12 = r"""Responsive material has been withheld pursuant to Tex. R. Civ. P. 193.3. The withheld material is responsive to this request and is withheld on the basis of applicable privileges."""
O13 = r"""Plaintiff withholds responsive material protected by the attorney-client privilege and provides this withholding statement pursuant to Rule 193.3."""
O14 = r"""Plaintiff objects and withholds responsive material to the extent it constitutes work product or material prepared in anticipation of litigation under TRCP 192.5."""
O15 = r"""Plaintiff objects to this request to the extent it seeks the identity, mental impressions, or opinions of consulting experts not expected to testify."""
O16 = r"""Plaintiff objects to the extent this request seeks confidential communications between spouses protected by the spousal communications privilege."""
O17 = r"""Plaintiff objects to this request to the extent it seeks highly sensitive mental-health information where Plaintiff has not affirmatively placed their mental condition at issue in this litigation."""
O18 = r"""Plaintiff objects to this request to the extent it seeks highly sensitive personal identifiers or private information without a demonstrated need that is proportional to the case."""
O19 = r"""Plaintiff objects to this interrogatory to the extent it requires a narrative response or detailed marshaling of proof more appropriately developed through deposition."""
O20 = r"""Plaintiff objects that this interrogatory improperly seeks to force Plaintiff to marshal all evidence supporting its claims or defenses."""
O21 = r"""Plaintiff objects to the extent this interrogatory requires a lay party Plaintiff to provide medical opinions beyond Plaintiff's personal knowledge or expert qualifications."""
O22 = r"""Plaintiff objects to this interrogatory to the extent it seeks expert information outside the scope or manner authorized by the expert discovery rules."""
O23 = r"""Plaintiff objects to the extent this interrogatory seeks information beyond Plaintiff's present knowledge and attempts to bind Plaintiff to a complete evidentiary statement before discovery is complete."""
O24 = r"""Plaintiff objects because this set exceeds the maximum number of permissible requests or answers under the governing TRCP discovery control plan."""
O25 = r"""Plaintiff objects to the extent this request seeks materials not within Plaintiff's possession, custody, or control."""
O26 = r"""Plaintiff objects because this request improperly requires Plaintiff to create a document that does not presently exist."""
O27 = r"""Plaintiff objects to signing the requested authorization in blank because it fails to specify the records sought or specific providers, depriving Plaintiff of a meaningful opportunity to evaluate relevance."""
O28 = r"""Plaintiff objects to the extent this request seeks a blanket medical authorization rather than relevant medical records properly subject to production or disclosure."""
O29 = r"""Plaintiff objects that this request seeks highly confidential tax information and is overbroad, intrusive, and not shown to be necessary in the form requested."""
O30 = r"""Plaintiff objects to this request to the extent it seeks Social Security records or identifiers that are irrelevant, overbroad, and unduly intrusive."""
O31 = r"""Plaintiff objects to this request to the extent it seeks personnel and employment records beyond those reasonably related to any claimed wage loss."""
O32 = r"""Plaintiff objects because this request seeks telephone records without a pleaded or otherwise demonstrated nexus to any claim or defense in the case."""
O33 = r"""Plaintiff objects to this request to the extent it seeks information regarding non-utilized health insurance, private health plans, or collateral benefits barred by the Texas Collateral Source Rule."""
O34 = r"""Plaintiff objects to the extent this request improperly expands the burden of medical billing disclosure by requiring the creation of an itemized analysis beyond the records themselves."""
O35 = r"""Plaintiff objects to the extent this request seeks a premature, exhaustive, or artificially fixed statement of damages before discovery is fully developed."""
O36 = r"""Plaintiff objects because this request for admission improperly seeks to establish a disputed merits issue rather than narrow an uncontroverted fact."""
O37 = r"""Plaintiff objects because this request for admission is compound and does not permit a fair admission or denial of a single proposition."""
O38 = r"""Plaintiff objects to this request for admission to the extent it seeks a pure legal conclusion."""
O39 = r"""After reasonable inquiry, Plaintiff lacks sufficient information to admit or deny this request and therefore denies it."""
O40 = r"""Plaintiff objects to the use of highly subjective or prejudicial terms within the request to the extent they assume a disputed characterization of the incident."""
O41 = r"""Plaintiff objects because the request for all written complaints across unrelated matters is overly broad, vague, and constitutes a prohibited fishing expedition."""
O42 = r"""Plaintiff objects to the extent the requested authorization would permit ex parte communications with healthcare providers rather than the production of defined records."""

# 2. Map directly into dictionary
TAXONOMY_OBJECTIONS = {
    "OB-TX-001: Relevance / Outside Scope": O1,
    "OB-TX-002: Overbroad as to Time": O2,
    "OB-TX-003: Overbroad as to Subject Matter": O3,
    "OB-TX-004: Vague / Ambiguous / Undefined Terms": O4,
    "OB-TX-005: Lack of Reasonable Particularity": O5,
    "OB-TX-006: Improper Fishing Expedition": O6,
    "OB-TX-007: Undue Burden / Expense": O7,
    "OB-TX-008: More Convenient / Less Burdensome Source":
