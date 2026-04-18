# Zettle Verkooprapport – de Fietsboot (Netlify)

Beheerpagina om het Zettle verkooprapport wekelijks bij te werken en te e-mailen.

## Stack
- **Netlify** (hosting + serverless functions)
- **Netlify Functions** (Node.js 18) — vervangt Azure Functions
- **Microsoft Graph API** — OneDrive upload + e-mail via M365
- **Zettle Purchase API v2** — aankopen ophalen

## Projectstructuur

```
zettle-rapport-netlify/
├── index.html
├── netlify.toml
└── netlify/
    └── functions/
        ├── package.json
        ├── shared/
        │   └── graph.js
        └── update-rapport.js      ← POST /api/update-rapport
```

## Deployment

### 1. Maak een Netlify site aan
```bash
# Installeer Netlify CLI (eenmalig)
npm install -g netlify-cli

# Log in
netlify login

# Initialiseer (vanuit de projectmap)
netlify init
```

Of koppel via de Netlify web interface aan een GitHub repo.

### 2. Installeer dependencies
```bash
cd netlify/functions
npm install
cd ../..
```

### 3. Stel omgevingsvariabelen in

Via Netlify Dashboard → Site settings → Environment variables, of via CLI:

```bash
netlify env:set GRAPH_TENANT_ID       "53ed2d57-347d-4bb4-bb4f-7a0473fb51fc"
netlify env:set GRAPH_CLIENT_ID       "9c30096f-5bcd-4044-855f-4c45b8ba90e7"
netlify env:set GRAPH_CLIENT_SECRET   "<secret>"
netlify env:set GRAPH_ONEDRIVE_USER   "admin@defietsboot.nl"
netlify env:set GRAPH_MAIL_FROM       "admin@defietsboot.nl"
netlify env:set ZETTLE_API_TOKEN      "<zettle-jwt-token>"
netlify env:set ZETTLE_RAPPORT_PATH   "MS365/Zettle Rapporten/Zettle_Verkooprapport_Actueel.xlsx"
```

### 4. Deploy
```bash
netlify deploy --prod
```

## Lokaal testen

Maak een `.env` bestand in de projectroot:
```env
GRAPH_TENANT_ID=53ed2d57-347d-4bb4-bb4f-7a0473fb51fc
GRAPH_CLIENT_ID=9c30096f-5bcd-4044-855f-4c45b8ba90e7
GRAPH_CLIENT_SECRET=<secret>
GRAPH_ONEDRIVE_USER=admin@defietsboot.nl
GRAPH_MAIL_FROM=admin@defietsboot.nl
ZETTLE_API_TOKEN=<token>
ZETTLE_RAPPORT_PATH=MS365/Zettle Rapporten/Zettle_Verkooprapport_Actueel.xlsx
```

Voeg `.env` toe aan `.gitignore`. Start de dev-server:
```bash
cd netlify/functions && npm install && cd ../..
netlify dev
```

## Omgevingsvariabelen

| Variabele | Beschrijving |
|---|---|
| `GRAPH_TENANT_ID` | Azure AD tenant ID |
| `GRAPH_CLIENT_ID` | App registration client ID |
| `GRAPH_CLIENT_SECRET` | App registration client secret |
| `GRAPH_ONEDRIVE_USER` | OneDrive eigenaar (`admin@defietsboot.nl`) |
| `GRAPH_MAIL_FROM` | Afzender e-mail (M365 mailbox) |
| `ZETTLE_API_TOKEN` | Zettle JWT access token |
| `ZETTLE_RAPPORT_PATH` | OneDrive pad + bestandsnaam voor het rapport |

## OneDrive map aanmaken

Zorg dat de map `MS365/Zettle Rapporten/` bestaat op OneDrive van `admin@defietsboot.nl`
voordat je de eerste keer deployt. Het bestand wordt aangemaakt/overschreven door de function.
