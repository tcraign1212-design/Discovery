# ──────────────────────────────────────────────
# 5. MAIN LAYOUT
# ──────────────────────────────────────────────
col1, col2 = st.columns([2, 3])

with col1:
    st.subheader("1. Case Details")
    case_stage = st.radio("Case Stage:", ["Pre-Litigation", "Active Litigation"], horizontal=True)
    ai_engine = st.radio("Model Engine:", ["Gemini (Google)", "ChatGPT (OpenAI)", "Claude (Anthropic)"], horizontal=True)

    # API Key Sieve
    if ai_engine == "Gemini (Google)":
        active_key = st.text_input("Gemini API Key", value=env_gemini_key or "", type="password", key="gem_key")
    elif ai_engine == "ChatGPT (OpenAI)":
        active_key = st.text_input("OpenAI API Key", value=env_openai_key or "", type="password", key="gpt_key")
    else:
        active_key = st.text_input("Anthropic API Key", value=env_anthropic_key or "", type="password", key="claude_key")

    st.markdown("---")
    case_type = st.selectbox("Framework:", ["Standard MVA", "Trucking", "Premises", "Workplace", "UM/UIM", "TTCA"])
    
    discovery_level = "N/A"
    if case_stage == "Active Litigation":
        discovery_level = st.radio("Discovery Level:", ["Level 1", "Level 2", "Level 3"], index=1, horizontal=True)

    st.markdown("**Risk Flags**")
    # 1. Government Entity Check
    government_entity = st.checkbox("Government Entity Involved")
    
    # 2. Universal Commercial Vehicle Check (Moved outside all IF blocks)
    comm_status = st.radio(
        "Commercial Vehicle Involved?", 
        ["No", "Yes", "Unsure"], 
        index=0, 
        horizontal=True,
        key="commercial_flag_permanent"  # Added key for session persistence
    )
    include_comm = (comm_status in ["Yes", "Unsure"])

    case_summary = st.text_area("Case Summary", height=160, placeholder="Detail the incident and liability theory...")
    
    with st.expander("Dates"):
        date_of_incident = st.text_input("Incident Date", placeholder="YYYY-MM-DD")
        sol_date = st.text_input("SOL Date", placeholder="YYYY-MM-DD")

    run_brief = st.button("Generate Case Intelligence Brief", type="primary", use_container_width=True)
