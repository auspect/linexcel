# crash_report — probe adversarial linexcel

Fixtures : `/tmp/kimi-wt/tests/fixtures/adversarial`
Généré par `tools/probe_crashes.py` — 18 fixtures sondées.

| fixture | phase | résultat | valeurs None | symptôme |
|---|---|---|---|---|
| `chaine_profonde.xlsx` | analyze | OK | 0/12 | analyse complète |
| `confidentialite.xlsx` | analyze | OK | 0/220 | analyse complète |
| `cycles.xlsx` | analyze | OK | 7/233 | 7/233 valeurs cellule None ; 1 warning(s) |
| `dates_1904.xlsx` | analyze | OK | 0/228 | analyse complète ; 1 warning(s) |
| `feuilles_unicode.xlsx` | analyze | OK | 0/9 | analyse complète |
| `fichier_tronque.xlsx` | analyze | EXCEPTION | — | ValueError: Could not analyze 'fichier_tronque.xlsx': File is not a zip file |
| `formules_croisees.xlsx` | analyze | OK | 0/233 | analyse complète |
| `iferror_nosheet.xlsx` | analyze | OK | 6/232 | 6/232 valeurs cellule None ; 2 warning(s) |
| `macro_hostile_corrompue.xlsm` | analyze | OK | 0/223 | analyse complète |
| `macro_security_block.xlsm` | analyze | OK | 0/223 | analyse complète |
| `macros_inoffensives.xlsm` | analyze | OK | 0/223 | analyse complète |
| `noms_definis_casses.xlsx` | analyze | OK | 1/224 | 1/224 valeurs cellule None |
| `pas_un_zip.xlsx` | analyze | EXCEPTION | — | ValueError: 'pas_un_zip.xlsx' is not an Excel file (xlsx/xlsm). Legacy .xls is not supported — re-save it as .xlsx first. |
| `realiste.xlsx` | analyze | OK | 0/220 | analyse complète |
| `refs_cassees_nosheet.xlsx` | analyze | OK | 4/228 | 4/228 valeurs cellule None ; 1 warning(s) |
| `refs_externes.xlsx` | analyze | OK | 2/229 | 2/229 valeurs cellule None ; 2 warning(s) |
| `sum_sur_formules.xlsx` | analyze | OK | 0/13 | analyse complète |
| `volatiles.xlsx` | analyze | OK | 2/228 | 2/228 valeurs cellule None |

## Synthèse

- exceptions/timeouts : **2**
- analyses OK avec valeurs None : **6**
- analyses propres : **10**

## Warnings remontés par l'analyse

- `cycles.xlsx` :
  - 5 cell(s) could not be computed by the engine and keep the value stored in the file: S1!G1, S1!F1, S1!L2, S1!F2, S1!F3
- `dates_1904.xlsx` :
  - S1!H1: recalculated 2022-08-30 differs from file value 2026-08-31
- `iferror_nosheet.xlsx` :
  - Global evaluation incomplete: #REF!: Sheet not found: NOSHEET
  - Values recovered cell by cell: 180 recomputed, 5 left to the value stored in the file
- `refs_cassees_nosheet.xlsx` :
  - Global evaluation completed after isolating 3 cell(s) whose references the engine cannot resolve; every other cell was recomputed. Only those keep the value stored in the file, if any. First blocker: #REF!: Sheet not found: NOSHEET
- `refs_externes.xlsx` :
  - Global evaluation completed after isolating 2 cell(s) whose references the engine cannot resolve; every other cell was recomputed. Only those keep the value stored in the file, if any. First blocker: #NAME?: Undefined name: '[Budget.xlsx]Annual'!B4
  - This workbook reads 1 external workbook(s). Neither read nor cached, so cells reading them have no value: Budget.xlsx. Pass refs_dir= (CLI: --refs-dir) to resolve them.
