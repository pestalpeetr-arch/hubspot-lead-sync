"""
HubSpot Lead Sync — Streamlit web app
Upload your lead Excel, connect to HubSpot, configure once, sync in one click.
"""
import streamlit as st
import requests
from hubspot_engine import (
    HubSpot, read_excel_bytes, determine_stage,
    count_active, run_sync,
)

# ─── Page setup ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="HubSpot Lead Sync",
    page_icon="🔄",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─── Session state defaults ───────────────────────────────────────────────────

for key in ["hs", "token", "pipelines", "stages", "last_pipeline_id",
            "companies", "sheets", "config", "sync_stats", "sync_log"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ─── Header ───────────────────────────────────────────────────────────────────

st.title("🔄 HubSpot Lead Sync")
st.caption(
    "Upload your lead Excel file, connect to HubSpot, pick your pipeline and stages, "
    "then hit **Sync**. Companies, contacts and deals are all created and linked automatically."
)

# ─── Step 1 · Upload & Connect ────────────────────────────────────────────────

st.divider()
st.subheader("Step 1 — Upload & Connect")

col_file, col_token = st.columns([1, 1])

with col_file:
    uploaded = st.file_uploader(
        "Lead Excel file (.xlsx)",
        type=["xlsx"],
        help="Your lead generation spreadsheet with the contact sheets",
    )

with col_token:
    token_input = st.text_input(
        "HubSpot Private App Token",
        type="password",
        placeholder="pat-eu1-...",
        help="HubSpot → Settings → Integrations → Private Apps",
    )

connect_disabled = not uploaded or not token_input
if st.button("🔗  Connect to HubSpot", use_container_width=True,
             disabled=connect_disabled):

    with st.spinner("Reading Excel and connecting to HubSpot…"):

        # ── Read Excel ────────────────────────────────────────────────────
        try:
            file_bytes = uploaded.read()
            companies, sheets = read_excel_bytes(file_bytes)
            if not companies:
                st.error("No contacts found. Make sure the file has the expected sheet structure.")
                st.stop()
        except Exception as e:
            st.error(f"Could not read Excel file: {e}")
            st.stop()

        # ── Connect to HubSpot ────────────────────────────────────────────
        try:
            hs = HubSpot(token_input)
            pipelines = hs.get_pipelines()
            if not pipelines:
                st.error("No deal pipelines found in your HubSpot account.")
                st.stop()
        except requests.HTTPError as e:
            code = e.response.status_code if e.response is not None else "?"
            if code == 401:
                st.error(
                    "❌ **Token invalid (401 Unauthorized).** "
                    "The token may have been rotated — get the latest one from "
                    "HubSpot → Settings → Integrations → Private Apps → your app → Auth tab."
                )
            elif code == 403:
                st.error(
                    "❌ **Missing scopes (403 Forbidden).** "
                    "Open your Private App in HubSpot and add **all** of these scopes, "
                    "then rotate the token and paste the new one here:\n\n"
                    "- `crm.schemas.deals.read`\n"
                    "- `crm.objects.companies.read` + `.write`\n"
                    "- `crm.objects.contacts.read` + `.write`\n"
                    "- `crm.objects.deals.read` + `.write`"
                )
            else:
                st.error(f"HubSpot API error: HTTP {code}. Check your token and try again.")
            st.stop()
        except Exception as e:
            st.error(f"Connection failed: {e}")
            st.stop()

        # ── Store in session ──────────────────────────────────────────────
        st.session_state.hs             = hs
        st.session_state.token          = token_input
        st.session_state.pipelines      = pipelines
        st.session_state.companies      = companies
        st.session_state.sheets         = sheets
        st.session_state.stages         = None
        st.session_state.last_pipeline_id = None
        st.session_state.config         = None
        st.session_state.sync_stats     = None
        st.session_state.sync_log       = None

# ─── Step 2 · Configure pipeline & stages ─────────────────────────────────────

if st.session_state.companies and st.session_state.pipelines:

    total, active, n_emails = count_active(st.session_state.companies)

    st.divider()

    # Summary banner
    col_a, col_b, col_c = st.columns(3)
    col_a.metric("Active companies", active,
                 help="Companies with at least one Contacted/Responded/Meeting = TRUE")
    col_b.metric("Total companies", total)
    col_c.metric("Contacts with email", n_emails)
    st.caption(f"📋 Sheets detected: **{', '.join(st.session_state.sheets)}**")

    st.divider()
    st.subheader("Step 2 — Pipeline & Stage Mapping")
    st.caption("Choose which HubSpot pipeline to use and map your three activity stages to deal stages.")

    # Pipeline selector
    pipeline_labels = [p["label"] for p in st.session_state.pipelines]
    chosen_label = st.selectbox("Pipeline", pipeline_labels)
    chosen_pipeline = next(
        p for p in st.session_state.pipelines if p["label"] == chosen_label
    )
    chosen_pipeline_id = chosen_pipeline["id"]

    # Fetch stages when pipeline changes
    if st.session_state.last_pipeline_id != chosen_pipeline_id:
        with st.spinner("Loading stages…"):
            try:
                stages = st.session_state.hs.get_pipeline_stages(chosen_pipeline_id)
                st.session_state.stages = stages
                st.session_state.last_pipeline_id = chosen_pipeline_id
            except Exception as e:
                st.error(f"Could not load stages: {e}")
                st.stop()

    if st.session_state.stages:
        stage_labels = [s["label"] for s in st.session_state.stages]
        stage_map_by_label = {s["label"]: s["id"] for s in st.session_state.stages}

        st.write("Map each activity status to the correct deal stage:")
        col1, col2, col3 = st.columns(3)
        with col1:
            s_contacted = st.selectbox("📧 Contacted →", stage_labels, key="sel_contacted")
        with col2:
            s_responded = st.selectbox("💬 Responded →", stage_labels, key="sel_responded")
        with col3:
            s_meeting = st.selectbox("📅 Meeting →", stage_labels, key="sel_meeting")

        st.session_state.config = {
            "pipeline_id": chosen_pipeline_id,
            "stage_map": {
                "contacted": stage_map_by_label[s_contacted],
                "responded": stage_map_by_label[s_responded],
                "meeting":   stage_map_by_label[s_meeting],
            },
        }

    # ─── Step 3 · Sync ────────────────────────────────────────────────────────

    if st.session_state.config:

        st.divider()
        st.subheader("Step 3 — Run Sync")
        st.caption(
            f"Will create/update **{active} companies**, their contacts, "
            f"and one deal per company in **{chosen_label}**."
        )

        if st.button("▶  Start Sync", use_container_width=True, type="primary"):
            # Reset previous results
            st.session_state.sync_stats = None
            st.session_state.sync_log   = []

            companies   = st.session_state.companies
            pipeline_id = st.session_state.config["pipeline_id"]
            stage_map   = st.session_state.config["stage_map"]
            hs          = st.session_state.hs

            total_active = sum(
                1 for c in companies.values() if determine_stage(c) is not None
            )

            log_placeholder = st.empty()
            progress_bar    = st.progress(0, text="Starting…")
            processed       = [0]
            log_lines       = []

            def log_callback(msg: str):
                log_lines.append(msg)
                st.session_state.sync_log = log_lines.copy()
                # Show last 40 lines so the box doesn't grow forever
                display = "\n".join(log_lines[-40:])
                log_placeholder.code(display, language=None)
                if msg.startswith("▸"):
                    processed[0] += 1
                    pct = min(processed[0] / max(total_active, 1), 1.0)
                    company_name = msg.split("▸")[-1].split("→")[0].strip()
                    progress_bar.progress(pct, text=f"Syncing {company_name}…")

            try:
                stats = run_sync(
                    companies=companies,
                    hs=hs,
                    pipeline_id=pipeline_id,
                    stage_map=stage_map,
                    log=log_callback,
                )
                progress_bar.progress(1.0, text="Done!")
                st.session_state.sync_stats = stats
            except requests.HTTPError as e:
                code = e.response.status_code if e.response else "?"
                st.error(f"Sync interrupted by API error: HTTP {code}. Check the log above.")
            except Exception as e:
                st.error(f"Sync failed unexpectedly: {e}")

        # ── Results ───────────────────────────────────────────────────────────

        if st.session_state.sync_log:
            if not st.session_state.sync_stats:
                # Log still visible during sync from the spinner above
                pass

        if st.session_state.sync_stats:
            stats = st.session_state.sync_stats
            st.divider()
            st.subheader("✅ Sync Complete")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Companies created", stats["co_created"])
            c2.metric("Contacts created",  stats["ct_created"])
            c3.metric("Deals created",     stats["deal_created"])
            c4.metric("Deals updated",     stats["deal_updated"])

            if stats["errors"]:
                st.warning(
                    f"⚠️  {stats['errors']} error(s) occurred — "
                    "check the log above for details."
                )
            if stats["skipped"]:
                st.info(
                    f"ℹ️  {stats['skipped']} companies skipped "
                    "(no Contacted / Responded / Meeting activity yet)."
                )

            if st.session_state.sync_log:
                with st.expander("📋 Full sync log"):
                    st.code("\n".join(st.session_state.sync_log), language=None)

# ─── Footer ───────────────────────────────────────────────────────────────────

st.divider()
st.caption(
    "Built for Rostra Group · "
    "Syncs companies, contacts & deals · "
    "All associations wired automatically"
)
