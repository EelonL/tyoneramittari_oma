import io
import json
import os
import time
import html
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font

# ============================================================
# Työnerämittari / Streamlit Community Cloud -yhteensopiva
# ------------------------------------------------------------
# Pääidea:
# 1) Käyttäjä syöttää työnerät.
# 2) Mittaus alkaa painamalla työnerää.
# 3) Työnerän vaihto tapahtuu painamalla seuraavaa työnerää.
# 4) Mittauksen lopetus tapahtuu painamalla punaista nappia.
# 5) Jokainen tapahtuma tallennetaan heti:
#    - st.session_stateen
#    - paikalliseen palautusjonoon (JSON)
# 6) Lopuksi data muunnetaan Exceliksi.
#
# HUOM:
# Tämä toteutus on tehty niin, että lyhyet verkkokatkokset eivät yleensä
# pilaa mittausta, jos selainvälilehti pysyy auki. Täydelliseen offline-
# kestävyyteen tarvitaan selaimen localStorage / PWA-ratkaisu tai natiiviappi.
# ============================================================

APP_TITLE = "Työnerämittari"
DATA_DIR = Path(".runtime_data")
DATA_DIR.mkdir(exist_ok=True)


def now_local() -> datetime:
    return datetime.now().astimezone()


def fmt_ts(dt: datetime) -> str:
    return dt.isoformat(timespec="seconds")


def parse_ts(value: str) -> datetime:
    return datetime.fromisoformat(value)


def duration_seconds(start_iso: str, end_iso: str) -> float:
    return (parse_ts(end_iso) - parse_ts(start_iso)).total_seconds()


def human_duration(seconds: float) -> str:
    seconds = int(round(seconds))
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def get_session_file() -> Path:
    if "session_file" not in st.session_state:
        session_id = f"session_{int(time.time() * 1000)}"
        st.session_state.session_file = str(DATA_DIR / f"{session_id}.json")
    return Path(st.session_state.session_file)


def default_state() -> Dict[str, Any]:
    return {
        "work_items": [],
        "events": [],
        "segments": [],
        "active_item": None,
        "active_start": None,
        "measurement_started": False,
        "measurement_finished": False,
        "measurement_label": "",
        "started_at": None,
        "finished_at": None,
        "upload_target": "Lataa Excel laitteelle",
        "onedrive_folder": "",
        "sharepoint_site": "",
        "sharepoint_drive": "",
        "sharepoint_folder": "",
    }


def ensure_state() -> None:
    if "app_state" not in st.session_state:
        st.session_state.app_state = default_state()
        recover_if_possible()


def persist_state() -> None:
    session_file = get_session_file()
    payload = st.session_state.app_state.copy()
    payload["last_saved_at"] = fmt_ts(now_local())
    session_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def recover_if_possible() -> None:
    # Palautetaan uusin sessiotiedosto, jos sellainen on olemassa.
    files = sorted(DATA_DIR.glob("session_*.json"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        return

    latest = files[0]
    try:
        recovered = json.loads(latest.read_text(encoding="utf-8"))
        if isinstance(recovered, dict) and recovered.get("events") is not None:
            st.session_state.app_state = {**default_state(), **recovered}
            st.session_state.session_file = str(latest)
    except Exception:
        pass


def add_event(event_type: str, item_name: Optional[str] = None, extra: Optional[Dict[str, Any]] = None) -> None:
    state = st.session_state.app_state
    ts = fmt_ts(now_local())
    payload = {"ts": ts, "type": event_type, "item": item_name}
    if extra:
        payload.update(extra)
    state["events"].append(payload)
    persist_state()


def start_item(item_name: str) -> None:
    state = st.session_state.app_state
    ts = fmt_ts(now_local())

    if not state["measurement_started"]:
        state["measurement_started"] = True
        state["started_at"] = ts

    if state["active_item"] is not None and state["active_start"] is not None:
        previous = {
            "item": state["active_item"],
            "start": state["active_start"],
            "end": ts,
            "duration_seconds": duration_seconds(state["active_start"], ts),
        }
        state["segments"].append(previous)
        add_event("switch", item_name, {"previous_item": state["active_item"]})
    else:
        add_event("start", item_name)

    state["active_item"] = item_name
    state["active_start"] = ts
    state["measurement_finished"] = False
    persist_state()


def stop_measurement() -> None:
    state = st.session_state.app_state
    if state["active_item"] is None or state["active_start"] is None:
        return

    ts = fmt_ts(now_local())
    segment = {
        "item": state["active_item"],
        "start": state["active_start"],
        "end": ts,
        "duration_seconds": duration_seconds(state["active_start"], ts),
    }
    state["segments"].append(segment)
    state["finished_at"] = ts
    state["measurement_finished"] = True
    add_event("stop", state["active_item"])
    state["active_item"] = None
    state["active_start"] = None
    persist_state()


def reset_all() -> None:
    st.session_state.app_state = default_state()
    persist_state()


def segments_df() -> pd.DataFrame:
    state = st.session_state.app_state
    rows: List[Dict[str, Any]] = []
    for i, seg in enumerate(state["segments"], start=1):
        rows.append(
            {
                "Järjestys": i,
                "Työnerä": seg["item"],
                "Alkuaika": seg["start"],
                "Loppuaika": seg["end"],
                "Kesto (s)": round(float(seg["duration_seconds"]), 1),
                "Kesto (hh:mm:ss)": human_duration(float(seg["duration_seconds"])),
            }
        )

    # Myös käynnissä oleva segmentti näkyviin esikatseluun
    if state["active_item"] and state["active_start"]:
        now_iso = fmt_ts(now_local())
        rows.append(
            {
                "Järjestys": len(rows) + 1,
                "Työnerä": state["active_item"],
                "Alkuaika": state["active_start"],
                "Loppuaika": "KÄYNNISSÄ",
                "Kesto (s)": round(float(duration_seconds(state["active_start"], now_iso)), 1),
                "Kesto (hh:mm:ss)": human_duration(float(duration_seconds(state["active_start"], now_iso))),
            }
        )

    return pd.DataFrame(rows)


def summary_df() -> pd.DataFrame:
    base = segments_df()
    if base.empty:
        return pd.DataFrame(columns=["Työnerä", "Toistot", "Yhteensä (s)", "Yhteensä (hh:mm:ss)"])

    finished_only = base[base["Loppuaika"] != "KÄYNNISSÄ"].copy()
    if finished_only.empty:
        return pd.DataFrame(columns=["Työnerä", "Toistot", "Yhteensä (s)", "Yhteensä (hh:mm:ss)"])

    grouped = (
        finished_only.groupby("Työnerä", as_index=False)
        .agg({"Kesto (s)": "sum", "Järjestys": "count"})
        .rename(columns={"Järjestys": "Toistot", "Kesto (s)": "Yhteensä (s)"})
    )
    grouped["Yhteensä (hh:mm:ss)"] = grouped["Yhteensä (s)"].apply(human_duration)
    grouped["Yhteensä (s)"] = grouped["Yhteensä (s)"].round(1)
    return grouped


def build_excel_bytes() -> bytes:
    detail = segments_df()
    summary = summary_df()
    meta = pd.DataFrame(
        [
            {"Kenttä": "Mittauksen nimi", "Arvo": st.session_state.app_state.get("measurement_label", "")},
            {"Kenttä": "Mittaus aloitettu", "Arvo": st.session_state.app_state.get("started_at", "")},
            {"Kenttä": "Mittaus lopetettu", "Arvo": st.session_state.app_state.get("finished_at", "")},
            {"Kenttä": "Aktiivinen sessiotiedosto", "Arvo": str(get_session_file())},
        ]
    )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        meta.to_excel(writer, index=False, sheet_name="Yhteenveto")
        start_row = len(meta) + 3
        summary.to_excel(writer, index=False, sheet_name="Yhteenveto", startrow=start_row)
        detail.to_excel(writer, index=False, sheet_name="Tapahtumat")

    output.seek(0)

    wb = load_workbook(output)
    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(value))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 40)

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output.getvalue()


def upload_to_m365_placeholder(file_bytes: bytes, filename: str) -> str:
    """
    Tähän kohtaan liitetään myöhemmin Microsoft Graph -tallennus.

    Toteutus vaatii käytännössä:
    - Entra ID / Azure App Registration
    - OAuth-kirjautumisen
    - Microsoft Graph -oikeudet
    - kohdekansion tunnistamisen (OneDrive/SharePoint)

    Tässä vaiheessa palautetaan informatiivinen viesti.
    """
    _ = file_bytes, filename
    return (
        "Valmis Graph-upload ei ole vielä kytketty tähän demo-versioon. "
        "Excel voidaan ladata laitteelle heti, ja tämän jälkeen tallentaa OneDriveen/Teamsiin."
    )


def save_ui() -> None:
    st.subheader("4. Excel-tiedoston muodostus ja tallennus")
    state = st.session_state.app_state

    filename = f"tyoneramittaus_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    excel_bytes = build_excel_bytes()

    target = st.radio(
        "Tallennustapa",
        ["Lataa Excel laitteelle", "OneDrive / Teams (Graph-integraatio)"],
        horizontal=False,
        key="save_target_radio",
    )
    state["upload_target"] = target
    persist_state()

    if target == "Lataa Excel laitteelle":
        st.download_button(
            "Lataa Excel",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption("Tämä toimii heti myös Streamlit Community Cloudissa.")
    else:
        st.text_input("OneDrive-kansio (esim. Mittaukset/Työmaa_A)", key="onedrive_folder_input")
        state["onedrive_folder"] = st.session_state.get("onedrive_folder_input", "")
        persist_state()

        if st.button("Tallenna OneDriveen / Teamsiin", use_container_width=True):
            msg = upload_to_m365_placeholder(excel_bytes, filename)
            st.info(msg)

        st.caption(
            "Teams-kanavien tiedostot sijaitsevat taustalla SharePointissa. "
            "Sama Graph-integraatio voidaan kohdistaa joko käyttäjän OneDriveen tai Teams/SharePoint-kirjastoon."
        )


def work_items_editor() -> None:
    st.subheader("1. Määritä työnerät")
    state = st.session_state.app_state

    state["measurement_label"] = st.text_input(
        "Mittauksen nimi",
        value=state.get("measurement_label", ""),
        placeholder="Esim. Huoneiston A123 työvaiheiden seuranta",
    )

    initial = "\n".join(state.get("work_items", []))
    text = st.text_area(
        "Syötä yksi työnerä per rivi",
        value=initial,
        height=180,
        placeholder="Esim.\nMuottityö\nRaudoitus\nBetonointi\nSiivous",
    )

    items = [row.strip() for row in text.splitlines() if row.strip()]
    state["work_items"] = items
    persist_state()

    cols = st.columns([1, 1])
    with cols[0]:
        if st.button("Tallenna työnerät", use_container_width=True):
            persist_state()
            st.success("Työnerät tallennettu.")
    with cols[1]:
        if st.button("Tyhjennä kaikki", use_container_width=True):
            reset_all()
            st.rerun()


def measurement_ui() -> None:
    st.subheader("2. Käynnistä mittaus")
    state = st.session_state.app_state

    if not state["work_items"]:
        st.warning("Lisää ensin vähintään yksi työnerä.")
        return

    st.caption(
        "Työnerät näkyvät yhdellä sivulla pystysuuntaisena listana. Ensimmäiset rivit näkyvät heti, ja lisää työneriä löytyy rullaamalla alaspäin."
    )

    st.markdown(
        """
        <style>
        :root {
            --tts-blue: #1973ff;
            --tts-blue-dark: #0f5fe0;
            --tts-blue-soft: #f3f7ff;
            --tts-border: #cfe0ff;
            --tts-text: #12324a;
            --tts-stop: #c62828;
        }
        div[data-testid="stButton"] > button {
            width: 100%;
            min-height: 78px;
            border-radius: 22px;
            text-align: left;
            font-size: 1.08rem;
            font-weight: 600;
            padding: 1rem 1.1rem;
            margin-bottom: 0.55rem;
            border: 1px solid var(--tts-border);
            background: var(--tts-blue-soft);
            color: var(--tts-text);
            box-shadow: 0 2px 8px rgba(25,115,255,0.08);
        }
        div[data-testid="stButton"] > button:hover {
            border-color: var(--tts-blue);
            box-shadow: 0 4px 12px rgba(25,115,255,0.14);
        }
        .active-work-item button {
            background: var(--tts-blue-dark) !important;
            color: white !important;
            border: 1px solid var(--tts-blue-dark) !important;
            min-height: 96px;
            box-shadow: 0 8px 18px rgba(25,115,255,0.28) !important;
        }
        .mobile-stop-bar {
            position: sticky;
            bottom: 0;
            background: rgba(255,255,255,0.98);
            padding-top: 0.45rem;
            padding-bottom: calc(0.45rem + env(safe-area-inset-bottom));
            border-top: 1px solid #e5e7eb;
            z-index: 9999;
            margin-top: 0.6rem;
        }
        .mobile-stop-bar div[data-testid="stButton"] > button {
            min-height: 68px;
            border-radius: 20px;
            background: var(--tts-stop) !important;
            color: white !important;
            border: 1px solid var(--tts-stop) !important;
            text-align: center;
            font-weight: 700;
            margin-bottom: 0;
        }
        .tts-app-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
            margin-bottom: 0.5rem;
            padding: 0.1rem 0 0.5rem 0;
        }
        .tts-app-title {
            font-size: 1.55rem;
            font-weight: 800;
            color: var(--tts-text);
            margin: 0;
            line-height: 1.15;
        }
        .tts-app-subtitle {
            font-size: 0.95rem;
            color: #5d7287;
            margin-top: 0.25rem;
        }
        .tts-logo-wrap {
            flex: 0 0 auto;
        }
        .tts-logo-wrap img {
            max-height: 52px;
            width: auto;
        }
        @media (max-width: 768px) {
            .block-container {
                padding-left: 0.45rem;
                padding-right: 0.45rem;
                padding-bottom: 5.5rem;
            }
            .tts-app-header {
                align-items: flex-start;
            }
            .tts-app-title {
                font-size: 1.28rem;
            }
            .tts-logo-wrap img {
                max-height: 42px;
            }
            div[data-testid="stButton"] > button {
                width: 100%;
                min-height: 84px;
                font-size: 1.08rem;
                padding: 1rem 0.95rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    active_item = state.get("active_item")
    active_start = state.get("active_start")
    active_elapsed = ""

    if active_item and active_start:
        elapsed_seconds = (now_local() - parse_ts(active_start)).total_seconds()
        active_elapsed = f"\n⏱ {human_duration(elapsed_seconds)}"

    for idx, item in enumerate(state["work_items"], start=1):
        is_active = active_item == item
        safe_item = html.escape(item)
        label = f"{idx}. {safe_item}"
        if is_active:
            label = f"🟢 {idx}. {safe_item}{active_elapsed}"
            st.markdown('<div class="active-work-item">', unsafe_allow_html=True)
            if st.button(label, key=f"item_{idx}", use_container_width=True):
                start_item(item)
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            if st.button(label, key=f"item_{idx}", use_container_width=True):
                start_item(item)
                st.rerun()

    if active_item:
        time.sleep(1)
        st.rerun()


def stop_button_ui() -> None:
    state = st.session_state.app_state

    st.markdown('<div class="mobile-stop-bar">', unsafe_allow_html=True)
    disabled = state["active_item"] is None
    stop_label = "🛑 Lopeta mittaus" if not disabled else "🛑 Ei aktiivista työnerää"
    if st.button(stop_label, use_container_width=True, disabled=disabled):
        stop_measurement()
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)


def recovery_info() -> None:
    with st.expander("Tietoa toimintavarmuudesta"):
        st.markdown(
            """
            - Jokainen aloitus, vaihto ja lopetus tallennetaan heti JSON-palautustiedostoon.
            - Lyhyt verkkokatkos ei yleensä riko mittausta, jos selainvälilehti pysyy auki.
            - Jos yhteys palaa, mittausta voidaan jatkaa ja Excel muodostaa lopuksi.
            - Täydellinen selainpäivitys tai sessioiden katkeaminen Streamlit Cloudissa voi silti katkaista tilan.
            - Tuotantoversiossa suosittelen lisäksi selaimen localStorage-varmistusta tai taustatietokantaa.
            """
        )


def live_tables() -> None:
    detail = segments_df()
    summary = summary_df()

    st.subheader("Mittausdata")
    t1, t2 = st.tabs(["Tapahtumat", "Yhteenveto"])
    with t1:
        if detail.empty:
            st.write("Ei vielä rivejä.")
        else:
            st.dataframe(detail, use_container_width=True)
    with t2:
        if summary.empty:
            st.write("Ei vielä yhteenvetoa.")
        else:
            st.dataframe(summary, use_container_width=True)


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="⏱️", layout="wide")
    ensure_state()

    logo_candidates = [
        "TTS_Logo_Blue_RGB_SA.jpg",
        "tts_logo.jpg",
        "tts_logo.png",
    ]
    logo_path = next((p for p in logo_candidates if Path(p).exists()), None)

    header_left, header_right = st.columns([4, 1])
    with header_left:
        st.markdown(
            """
            <div class="tts-app-header">
                <div>
                    <div class="tts-app-title">⏱️ Työnerämittari</div>
                    <div class="tts-app-subtitle">Työnerien käynnistys, vaihto ja lopetus yhdellä näkymällä</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with header_right:
        if logo_path:
            st.image(logo_path, use_container_width=True)

    st.write(
        "Tällä sovelluksella voit mitata työnerien alkamis- ja päättymisaikoja sekä muodostaa lopuksi Excel-tiedoston."
    )

    work_items_editor()
    st.divider()
    measurement_ui()
    stop_button_ui()
    recovery_info()
    st.divider()
    live_tables()
    st.divider()
    save_ui()


if __name__ == "__main__":
    main()
