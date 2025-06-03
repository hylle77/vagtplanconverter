import streamlit as st
import re
from io import BytesIO
from datetime import datetime, timedelta
from ics import Calendar, Event
import docx

# --- Sideopsætning ---
st.set_page_config(
    page_title="DOCX (WORD) til Kalender (.ics)",
    layout="centered",
    initial_sidebar_state="expanded",
)

# Hovedtitel
st.title("🗓️ Konverter Vagtplan til Kalender")

# Korte instruktioner øverst
st.markdown(
    """
    Hej! Her kan du uploade din **vagtplan (.docx)** og få en **.ics-fil**, som du kan importere til din kalender.
    __________________________________________________________
    Følg trinene nedenfor:
    1. Upload din vagtplan i `.docx`-format (WORD).
    2. Vælg dit navn fra listen (eller indtast det manuelt).
    3. Angiv en titel for dine kalenderbegivenheder.
    4. Download din `.ics`-fil, når vagterne er fundet.
    """
)

# ---- Hjælpefunktioner ----

def parse_docx(file):
    """Læs en .docx-fil og returner al teksten."""
    doc = docx.Document(file)
    return "\n".join(p.text for p in doc.paragraphs)

def normalize_time_str(ts):
    """
    Omformater “13.30” → “13:30”; hvis der ikke er “:”, tilføj “:00”; derefter valider.
    Returnerer “HH:MM”. Kaster ValueError ved ugyldigt format.
    """
    s = ts.strip().lower().replace(".", ":")
    if s == "luk":
        return s
    if ":" not in s:
        s += ":00"
    try:
        h, m = map(int, s.split(":"))
        if not (0 <= h < 24 and 0 <= m < 60):
            raise ValueError
    except ValueError:
        raise ValueError(f"Ugyldigt tidsformat: '{ts}' (efter normalisering: '{s}')")
    return f"{h:02d}:{m:02d}"

def get_weekday_closing(end_raw, date_str, closing_hours):
    """
    Hvis end_raw er 'luk' eller '??', brug closing_hours[weekday] for man–søn.
    Ellers normaliser via normalize_time_str().
    """
    date_obj = datetime.strptime(date_str, "%d/%m/%Y")
    wd = date_obj.weekday()  # 0=Mandag … 6=Søndag
    s = end_raw.strip().lower()

    if s in ("luk", "??"):
        close_hour = closing_hours[wd]
        return "02:00" if close_hour < 6 else f"{close_hour:02d}:00"

    return normalize_time_str(end_raw)

def normalize_name(name):
    """
    Fjern alt, der ikke er A–Z/æøå/ÆØÅ/mellemrum/.-, strip evt. ord som rengør, rengøring, clean, opvask osv.
    Returner lowercased.
    """
    cleaned = re.sub(r"[^a-zA-ZæøåÆØÅ\s\.\-]", "", name).strip()
    cleaned = re.sub(
        r"\b(rengør|rengøring|clean|opvask|vask|støvsug|note).*",
        "",
        cleaned,
        flags=re.IGNORECASE,
    ).strip()
    return cleaned.lower()

def extract_shifts(text, target_names, year, closing_hours):
    """
    Gennemgå tekstlinjer og udtræk events for skift, der matcher navnene i target_names.
    """
    pattern_date1 = re.compile(r"\b\w+\.?\s*d\.?\s*(\d{1,2})/(\d{1,2})", re.IGNORECASE)
    pattern_date2 = re.compile(r"^(.+?):\s*(\d{1,2})/(\d{1,2})", re.IGNORECASE)
    pattern_shift = re.compile(
        r"(\d{1,2}(?:[:\.]\d{2})?)[\s\-–]+(\d{1,2}(?:[:\.]\d{2})?|luk|\?\?)\s*(?::\s*|\s{2,})(.+)",
        re.IGNORECASE,
    )

    events = []
    current_date = None
    custom_location = None
    in_moments = False
    norm_targets = [normalize_name(t) for t in target_names]

    for line in text.splitlines():
        raw = line.strip()
        if not raw:
            in_moments = False
            continue

        # Fri tekst “Navn: DD/MM”
        dm2 = pattern_date2.match(raw)
        if dm2:
            prefix, d, m = dm2.groups()
            current_date = f"{int(d):02d}/{int(m):02d}/{year}"
            custom_location = prefix.strip()
            in_moments = "moments" in prefix.lower()
            continue

        # Ugedag header “Mandag d.DD/MM”
        dm1 = pattern_date1.search(raw)
        if dm1:
            d, m = dm1.groups()
            current_date = f"{int(d):02d}/{int(m):02d}/{year}"
            custom_location = None
            in_moments = False
            continue

        if current_date is None:
            continue

        sm = pattern_shift.match(raw)
        if not sm:
            continue

        start_raw, end_raw, names_line = sm.groups()
        try:
            start_ts = normalize_time_str(start_raw)
            end_ts = get_weekday_closing(end_raw, current_date, closing_hours)
        except Exception:
            continue

        try:
            dt_start = datetime.strptime(f"{current_date} {start_ts}", "%d/%m/%Y %H:%M")
            dt_end   = datetime.strptime(f"{current_date} {end_ts}",   "%d/%m/%Y %H:%M")
            if dt_end <= dt_start:
                dt_end += timedelta(days=1)
        except Exception:
            continue

        raw_desc = "Moments" if in_moments else "Bar"

        for nm in re.split(r",| og ", names_line):
            note_match = re.search(r"\(\s*(rengør|rengøring|støvsug|opvask|vask|clean|note)\s*\)", nm, re.IGNORECASE)
            note = note_match.group(1).lower() if note_match else None

            cleaned = normalize_name(nm)
            if cleaned in norm_targets:
                if custom_location:
                    location = custom_location
                else:
                    place = "Moments" if in_moments else "Bar"
                    location = f"D'Wine {place}, Algade 54, 9000 Aalborg"

                full_desc = raw_desc + (f" – {note}" if note else "")
                events.append({
                    "name": cleaned,
                    "start": dt_start,
                    "end": dt_end,
                    "raw": full_desc,
                    "location": location
                })

    return events

def create_ics(events, custom_title):
    """Opret en .ics-fil i hukommelsen baseret på events."""
    cal = Calendar()
    for e in events:
        ev = Event()
        ev.name = custom_title
        ev.begin = e["start"]
        ev.end = e["end"]
        ev.description = e["raw"]
        ev.location = e["location"]
        cal.events.add(ev)

    output = BytesIO()
    output.write(str(cal).encode("utf-8"))
    output.seek(0)
    return output

# ---- Streamlit-UI ----

# Fast årstal
ÅR = 2025

# Upload-knap og titelinput
uploaded_file = st.file_uploader("📂 Vælg vagtplan (.docx)", type=["docx"])
custom_title = st.text_input("Skriv titel til kalenderbegivenheder", value="🤓 - Arbejde")

if uploaded_file:
    with st.spinner("Læser fil og udtrækker vagter…"):
        raw_text = parse_docx(uploaded_file)

    # ---- Find alle navne i vagtplanen ----
    all_names_norm = []
    name_pattern = re.compile(
        r"(\d{1,2}(?:[:\.]\d{2})?)[\s\-–]+(\d{1,2}(?:[:\.]\d{2})?|luk|\?\?)\s*(?::\s*|\s{2,})(.+)",
        re.IGNORECASE,
    )
    for line in raw_text.splitlines():
        m = name_pattern.match(line.strip())
        if not m:
            continue
        _, _, names_line = m.groups()
        for nm in re.split(r",| og ", names_line):
            cleaned = normalize_name(nm)
            if cleaned and cleaned not in all_names_norm:
                all_names_norm.append(cleaned)
    all_names_norm.sort()

    # ---- Vis navne i en dropdown (eller mulighed for manuel indtastning) ----
    if all_names_norm:
        st.subheader("Vælg dit navn")
        all_names_display = [n.title() for n in all_names_norm]
        selected_display = st.selectbox("", all_names_display)
        selected_name = selected_display.lower()
    else:
        st.subheader("Indtast dit navn manuelt")
        typed = st.text_input("Skriv dit navn her", placeholder="F.eks. Lasse Hansen")
        selected_name = normalize_name(typed) if typed else ""

    # Når brugeren har valgt/indtastet navn
    if selected_name:
        with st.spinner("Finder vagter for dig…"):
            shifts = extract_shifts(
                raw_text,
                [selected_name],
                ÅR,
                closing_hours=[22, 22, 23, 23, 2, 2, 22],  # Man→Søn
            )

        if not shifts:
            st.warning("🕵️‍♂️ Ingen vagter fundet for det valgte navn. Tjek du har stavet korrekt eller prøv et andet navn.")
        else:
            st.success(f"✅ Fandt {len(shifts)} vagt(er) for '{selected_display if all_names_norm else typed.title()}'.")
            ics_file = create_ics(shifts, custom_title)

            # Download-knap
            st.download_button(
                label="📥 Download din .ics-fil",
                data=ics_file,
                file_name="vagtplan.ics",
                mime="text/calendar",
            )

            # Balloner
            st.balloons()

            # Ekstra info
            st.markdown(
                """
                **Tip:** Importér `vagtplan.ics` i din foretrukne kalender-app (Google Kalender, Outlook, Apple Kalender osv.), helt automatisk, hvis i er ligeså dovne som mig! 🎉
                """
            )
