# Demomanus – Spårbar (cirka 10 minuter)

**Budskapet i en mening:** Rapporterna är redan uträknade när halvåret är slut. Du tittar på det som är markerat,
skriver en rad till kunden och godkänner. Klart.

Mötet är för henne som gör rapporterna i dag. Visa vyn **Enkel** först och håll det enkelt. Prata om
"verktyget", inte om tekniken. Frågar hon hur texten skrivs: säg som det är, att en AI-modell skriver ett förslag
utifrån siffrorna och att hon alltid läser igenom det innan något skickas.

## Före mötet

- [ ] `npm start` och kontrollera att terminalen visar **AI: live**
- [ ] Öppna **http://127.0.0.1:3000/?reset** i helskärm. Det nollställer demon: inga hanterade kontroller, inga
      låsta rapporter, vyn **Enkel**, svenska. Portalen öppnar på rapporten för första kunden
- [ ] Ha **ett foto av ett tankkvitto** redo på skrivbordet (och **Milj.xlsx** om chefen är med)
- [ ] Fråga MLT om det är okej att verktyget används på deras egen fil (bara summerade siffror skickas)

## 1. Rapporten (1 min)

- "Så här ser det ut morgonen efter att halvåret är slut. Rapporten för första kunden är redan uträknad, direkt från
  affärssystemet. Du kopierar ingenting."
- Peka på de tre stegen till vänster och kundens rapport till höger. "Det här är fiktiva kunder."

## 2. Steg 1: Kontrollera (3 min)

- Fyra saker är markerade. Klicka **Vad betyder det?** på tonnaget: förklaringen kommer på vanlig svenska.
- Klicka **Det stämmer** på *Interna fordon*, välj ett förslag och **Spara**. "Det här gäller alla kunder, så du
  svarar bara en gång."
- Hantera resten på samma sätt. Rapporten till höger är alltid aktuell.

## 3. Steg 2 och 3: Text och godkänn (2 min)

- Klicka **Skriv texten åt mig**. Ändra ett ord. "Siffrorna i tabellerna rör den aldrig."
- Byt rapportspråk till **English** och tillbaka: kunden kan få rapporten på engelska.
- Klicka **Godkänn och lås** och bekräfta. "Nu är den låst med siffror, dina kommentarer och en verifieringskod.
  PDF:en kan gå till kunden."
- Klicka **Nästa kund**: Demokund B har bara en sak kvar, eftersom svaren som gäller alla kunder redan är gjorda.
- Visa **Alla kunder**: "1 av 3 klara".

## 4. Kvitton (1 min)

- Gå till **Kvitton** och släpp fotot av tankkvittot. Datum, drivmedel, liter och belopp fylls i.
- "Ingenting sparas förrän du har tittat på det."

## 5. Vyn Avancerad (1 min, om chefen är med)

- Byt till **Avancerad** längst ned till vänster: hela portalen med översikt, faktorer, regelverk och AI-assistent.
- **Godkända** / **Arkiv**: öppna den låsta rapporten och visa kommentarerna och verifieringskoden.
- Valfritt: **Data** och släpp **Milj.xlsx** → **Översikt** visar **Matchar Excel** (25,1 ton CO₂e).

## 6. Avslut (1 min)

- "I den färdiga versionen körs allt i er egen Microsoft-miljö. Vi på Brick hanterar ingen av er data."
- Fråga om de sista 10 procenten av underlaget: var finns drivmedelspriset, hållbarhetsbesked för HVO,
  avstånd per order och uppgifter från inhyrda åkerier?

## Om något strular

- **Svaret dröjer:** fortsätt prata, det brukar komma inom 5–15 sekunder.
- **Ingen internetanslutning:** starta utan nyckel (tom `ANTHROPIC_API_KEY`). Allt fungerar; texterna är då
  märkta som demotexter.
- **Fel API-nyckel:** felet visas i portalen; kontrollera `.env` och starta om.
- **Börja om:** öppna http://127.0.0.1:3000/?reset.
