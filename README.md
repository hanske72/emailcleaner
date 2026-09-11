- lese, skrive eller slette `/root/.ssh`, `~/.ssh` eller `authorized_keys`
- endre `/etc/ssh`, SSH-porten, `ssh`/`sshd`-tjenesten, UFW eller CrowdSec

SSH-tilgang og serverherding skal alltid behandles som en separat, manuelt godkjent oppgave. Etter publisering skal en **ny** SSH-tilkobling til port 7222 verifiseres før arbeidet regnes som ferdig. Hvis den feiler, skal dette meldes som en kritisk publiseringsfeil selv om nettsiden fortsatt svarer.
# E-postrydder for Microsoft 365 Outlook

Dette er det aktive prosjektet for E-postrydder. Gmail-klienten er avviklet; videre utvikling gjelder bare Microsoft 365 Outlook.

## Arkitektur

- Office Add-in med React og TypeScript
- Microsoft Graph for postbokser, mapper og meldinger
- Office Dialog API for Microsoft-innlogging
- Google Gemini for AI-analyse (AI-leverandøren er uavhengig av den avviklede Gmail-klienten)
- Webpack bygger statiske filer til `dist/`
- `manifest.xml` peker Outlook mot `https://geolead.no/addin/taskpane.html`

Aktiv kildekode ligger i `src/`. Mappen `dist/` er generert og skal ikke versjonskontrolleres.

## Lokal utvikling

```powershell
npm ci
npm run validate
npm run build
```

## Versjonsstatus per 10. september 2026

- GitHub `main`: første versjon, applikasjonsversjon 4.1
- Offentlig live-build under `/addin/`: applikasjonsversjon 4.1
- Lokal arbeidskopi: ukommitterte stabilitets- og loggendringer merket 4.6
- Office-manifest: 2.6.0.0 (egen manifestversjon)

De lokale 4.6-endringene må testes og committes før de kan regnes som gjeldende kildekode. De skal ikke publiseres automatisk.

## Kritisk publiseringsregel

GeoLead og E-postrydder deler server og domene. En publisering av E-postrydder skal derfor aldri:

- publisere eller synkronisere hele GeoLead-webroten
- bruke sletting eller speiling mot en overordnet servermappe
- endre Nginx-konfigurasjonen
- restarte eller erstatte GeoLead-applikasjonen
- lese, skrive eller slette `/root/.ssh`, `~/.ssh` eller `authorized_keys`
- endre `/etc/ssh`, SSH-porten, `ssh`/`sshd`-tjenesten, UFW eller CrowdSec

Bare de ferdig bygde filene fra `dist/` skal kopieres til serverens dedikerte mappe `/var/www/geolead/wwwroot/addin/`. Eksisterende live-filer skal sikkerhetskopieres først. SSH går på port 7222; bruk tilkoblingsoppsettet fra GeoLead-prosjektet og aldri anta standardport eller en gammel nøkkel.

SSH-tilgang og serverherding skal alltid behandles som en separat, manuelt godkjent oppgave. Etter publisering skal en **ny** SSH-tilkobling til port 7222 verifiseres før arbeidet regnes som ferdig. Hvis den feiler, skal dette meldes som en kritisk publiseringsfeil selv om nettsiden fortsatt svarer.

Etter hver publisering skal minst disse endepunktene kontrolleres:

1. `https://geolead.no/`
2. `https://geolead.no/Leads/Compose`
3. `https://geolead.no/addin/taskpane.html`
4. `https://geolead.no/addin/taskpane.js`

GeoLeads release-workflow bevarer `wwwroot/addin` mellom releaser. Ikke endre denne persistensregelen.