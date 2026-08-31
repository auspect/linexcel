# crash_report — probe adversarial linexcel

Fixtures : `/tmp/kimi-wt/tests/fixtures/adversarial`
Généré par `tools/probe_crashes.py` — 11 fixtures sondées.

| fixture | phase | résultat | valeurs None | symptôme |
|---|---|---|---|---|
| `confidentialite.xlsx` | analyze | OK | 0/220 | analyse complète |
| `cycles.xlsx` | analyze | OK | 7/233 | 7/233 valeurs cellule None |
| `dates_1904.xlsx` | analyze | OK | 0/228 | analyse complète |
| `formules_croisees.xlsx` | analyze | OK | 0/233 | analyse complète |
| `iferror_nosheet.xlsx` | analyze | OK | 6/232 | 6/232 valeurs cellule None |
| `macro_hostile_corrompue.xlsm` | analyze | OK | 0/223 | analyse complète |
| `macro_security_block.xlsm` | analyze | OK | 0/223 | analyse complète |
| `macros_inoffensives.xlsm` | analyze | OK | 0/223 | analyse complète |
| `realiste.xlsx` | analyze | OK | 0/220 | analyse complète |
| `refs_cassees_nosheet.xlsx` | analyze | OK | 4/228 | 4/228 valeurs cellule None |
| `sum_sur_formules.xlsx` | analyze | OK | 0/13 | analyse complète |

## Synthèse

- exceptions/timeouts : **0**
- analyses OK avec valeurs None : **3**
- analyses propres : **8**
