# Demomanus – Spårbar (cirka 10 minuter)

**Budskapet i en mening:** AI läser, skriver och förklarar. Siffrorna kommer från en fast beräkning som matchar er
Excel till decimalen, och en människa godkänner varje rapport.

## Före mötet

- [ ] `npm start` och kontrollera att terminalen visar **AI: live**
- [ ] Öppna http://127.0.0.1:3000 i helskärm
- [ ] Ha **Milj.xlsx** (MLT:s egen fil) och **ett foto av ett tankkvitto** redo på skrivbordet
- [ ] Fråga MLT om det är okej att AI-funktionerna används på deras egen fil (bara summerade siffror skickas)

## 1. Översikt (1 min)

- "Det här är en fiktiv demokund. Allt ni ser räknas fram från orderraderna."
- Klicka på stapeln **Lastväxlare 3-axl**: "Varje siffra går att spåra ner till orderraden och formeln."

## 2. Er egen fil (1 min)

- Gå till **Data** och släpp **Milj.xlsx**.
- Gå till **Översikt**: 25,1 ton CO₂e och märket **Matchar Excel**.
- "Samma resultat som er Excel, utan att någon kopierar in något."

## 3. Kontroller (1,5 min)

- Gå till **Kontroller**. Två allvarliga avvikelser som Excel aldrig visar, till exempel 2 682 ton från rader som inte är fakturerade.
- Klicka **Förklara med AI** på tonnagekontrollen. Claude förklarar på vanlig svenska och föreslår åtgärder.

## 4. Kundrapport (2 min)

- Gå till **Kundrapport** och klicka **Skriv sammanfattning med AI**. Texten skrivs medan ni tittar.
- "Texten går att redigera. Siffrorna i tabellerna rör AI:n aldrig."
- Byt rapportspråk till **English**, skriv om, och visa att kunden kan få rapporten på engelska.
- Klicka **PDF** och öppna filen.

## 5. AI-assistent (1,5 min)

- Gå till **AI-assistent** och klicka på frågorna:
  - "Vilka fordonstyper står för mest utsläpp?"
  - "Vad händer om vi använder svenskt marknadssnitt för HVO?"
- Låt någon från MLT ställa en egen fråga.

## 6. Kvitton (1 min)

- Gå till **Kvitton** och släpp fotot av tankkvittot. Claude läser datum, drivmedel, liter och belopp.
- "Ingenting sparas förrän en människa har granskat raderna."

## 7. Avslut (1 min)

- "I produktion körs allt i er egen Azure och Microsoft 365. Vi på Brick hanterar ingen av er data."
- "Rapporterna tas fram automatiskt varje halvår, ni granskar och godkänner, och kunden får dem per e-post."
- Nästa steg: beslutslistan och checklistan i readiness-paketet.

## Om något strular

- **AI-svaret dröjer:** fortsätt prata, det brukar komma inom 5–15 sekunder.
- **Ingen internetanslutning:** starta utan nyckel (tom `ANTHROPIC_API_KEY`). Allt fungerar, AI-svaren är då märkta som demotexter.
- **Fel API-nyckel:** felet visas i portalen; kontrollera `.env` och starta om.
