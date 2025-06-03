# DOCX til Kalender (.ics)

📆 En enkel Streamlit-app til at omdanne vagtplaner i `.docx`-format til en `.ics`-kalenderfil

Denne applikation kan bruges af alle, der ønsker at samle en arbejdsplan i én digital kalender. Upload din vagtplan i `.docx`-format, vælg det relevante navn og få en `.ics`-fil klar til import i Google Kalender, Outlook, Apple Kalender eller andre kalenderprogrammer.

---

## Funktioner

- **Upload af `.docx`-vagtplan**  
  Træk eller vælg en fil med arbejdsplanen. App’en udtrækker automatisk al relevant tekst.

- **Automatisk navneudtrækning**  
  Alle navne, der optræder i filen ved vagter, opsamles og vises i en dropdown-menu. Du kan også indtaste dit navn manuelt, hvis det ikke dukker op.

- **Tidsnormalisering**  
  Formater som `13.30` omformes til `13:30`. Hvis sluttidspunkt er tidligere end starttidspunkt, flyttes sluttiden til næste dag.

- **Fleksible datooverskrifter**  
  Understøtter overskrifter som “Mandag d. 4/4” eller “Afdeling: 5/4” til at sætte dato for de følgende vagter. Vagterne grupperes korrekt under disse datooverskrifter.

- **Parentetiske bemærkninger**  
  Noter i parentes (f.eks. `(rengøring)`, `(opvask)`) medtages i beskrivelsen af hver vagt.

- **Generering af `.ics`**  
  Alle fundne vagter pakkes i en iCalendar-fil, der kan downloades og importeres i enhver kalender-app.

---
