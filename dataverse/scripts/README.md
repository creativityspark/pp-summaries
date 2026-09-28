# Dataverse scripts

| Script | Purpose |
| --- | --- |
| `inventory-solution.mjs` (`npm run inventory`) | Read-only inventory of the `BizzSummit2026` solution. |
| `migrate-summary-studio.mjs` (`npm run migrate`) | Creates the Summary Studio data model and seeds the English demo data (recipe, prompt, three accounts, activities, summaries, runs). |

Both scripts authenticate with the device-code flow against the Spark Tools DEV environment. Earlier prototype scripts were removed; they remain available in the git history.
