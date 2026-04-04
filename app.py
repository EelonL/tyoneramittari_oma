import io
import json
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font

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
    seconds = max(0, int(round(seconds)))
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
    }


def recover_if_possible() -> None:
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


def ensure_state() -> None:
    if "app_state" not in st.session_state:
        st.session_state.app_state = default_state()
        recover_if_possible()


def persist_state() -> None:
    session_file = get_session_file()
    payload = dict(st.session_state.app_state)
    payload["last_saved_at"] = fmt_ts(now_local())
    session_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def add_event(event_type: str, item_name: Optional[str] = None, extra: Optional[Dict[str, Any]] = None) -> None:
    state = st.session_state.app_state
    payload = {
        "ts": fmt_ts(now_local()),
        "type": event_type,
        "item": item_name,
    }
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


def render_brand_header() -> None:
    logo_candidates = [
        "TTS_Logo_Blue_RGB_SA.jpg",
        "tts_logo.jpg",
        "tts_logo.png",
    ]
    logo_path = next((p for p in logo_candidates if Path(p).exists()), None)

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
            margin-bottom: 0.75rem;
        }

        .work-items-section div[data-testid="stButton"] > button {
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

        .work-items-section div[data-testid="stButton"] > button:hover {
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

        .stop-section div[data-testid="stButton"] > button {
            min-height: 72px;
            border-radius: 20px;
            background: var(--tts-stop) !important;
            color: white !important;
            border: 1px solid var(--tts-stop) !important;
            text-align: center;
            font-weight: 700;
            margin-top: 0.25rem;
        }

        @media (max-width: 768px) {
            .block-container {
                padding-left: 0.45rem;
                padding-right: 0.45rem;
                padding-bottom: 2rem;
            }

            .tts-app-title {
                font-size: 1.28rem;
            }

            .work-items-section div[data-testid="stButton"] > button {
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

    c1, c2 = st.columns([4, 1])
    with c1:
        st.markdown(
            """
            <div class="tts-app-title">⏱️ Työnerämittari</div>
            <div class="tts-app-subtitle">Työnerien käynnistys, vaihto ja lopetus yhdellä näkymällä</div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        if logo_path:
            st.image(logo_path, use_container_width=True)

    st.write("Tällä sovelluksella voit mitata työnerien alkamis- ja päättymisaikoja sekä muodostaa lopuksi Excel-tiedoston.")


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
        height=160,
        placeholder="Esim.\nMuottityö\nRaudoitus\nBetonointi\nSiivous",
    )

    items = [row.strip() for row in text.splitlines() if row.strip()]
    state["work_items"] = items
    persist_state()

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Tallenna työnerät", use_container_width=True):
            persist_state()
            st.success("Työnerät tallennettu.")
    with c2:
        if st.button("Tyhjennä kaikki", use_container_width=True):
            reset_all()
            st.rerun()


@st.fragment(run_every="1s")
def measurement_ui() -> None:
    st.subheader("2. Käynnistä mittaus")
    state = st.session_state.app_state

    if not state["work_items"]:
        st.warning("Lisää ensin vähintään yksi työnerä.")
        return

    st.caption("Työnerät näkyvät yhdellä sivulla pystysuuntaisena listana. Lisää työneriä löytyy rullaamalla alaspäin.")

    active_item = state.get("active_item")
    active_start = state.get("active_start")
    active_elapsed = ""

    if active_item and active_start:
        elapsed_seconds = (now_local() - parse_ts(active_start)).total_seconds()
        active_elapsed = f" ⏱ {human_duration(elapsed_seconds)}"

    st.markdown('<div class="work-items-section">', unsafe_allow_html=True)

    for idx, item in enumerate(state["work_items"], start=1):
        is_active = active_item == item
        label = f"{idx}. {item}"

        if is_active:
            label = f"🟢 {idx}. {item}{active_elapsed}"
            st.markdown('<div class="active-work-item">', unsafe_allow_html=True)
            if st.button(label, key=f"item_{idx}", use_container_width=True):
                start_item(item)
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
        else:
            if st.button(label, key=f"item_{idx}", use_container_width=True):
                start_item(item)
                st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


def stop_button_ui() -> None:
    state = st.session_state.app_state
    active_item = state.get("active_item")
    active_start = state.get("active_start")

    st.markdown("---")
    st.subheader("3. Mittauksen lopetus")
    st.caption("Lopeta käynnissä oleva työnerä tästä painikkeesta.")

    stop_label = "🛑 Lopeta mittaus"
    if active_item:
        stop_label = f"🛑 Lopeta: {active_item}"

    st.markdown('<div class="stop-section">', unsafe_allow_html=True)
    if st.button(stop_label, key="stop_measurement_main", use_container_width=True):
        if active_item and active_start:
            stop_measurement()
            st.rerun()
        else:
            st.warning("Aktiivista työnerää ei ole käynnissä.")
    st.markdown("</div>", unsafe_allow_html=True)


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


def save_ui() -> None:
    st.subheader("4. Excel-tiedoston muodostus")
    filename = f"tyoneramittaus_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    excel_bytes = build_excel_bytes()

    st.download_button(
        "Lataa Excel",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Excel ladataan käyttäjän laitteelle.")


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="⏱️", layout="wide")
    ensure_state()

    render_brand_header()
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
