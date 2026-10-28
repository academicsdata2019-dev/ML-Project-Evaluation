# Semester 7 ML Project Evaluation Portal

Complete Streamlit project package with the application, SQLite database, student roster, requirements, and 22 local project illustrations.

## Included files
- `app.py` — Streamlit evaluation portal with responsive project-group tiles, concise one-sentence project descriptions, and a consistent, non-cropped page heading on every authenticated page.
- `ml_project_evaluation.db` — SQLite database from the current project package. Existing student records and evaluation data are preserved.
- `student_roster.csv` — student roster.
- `requirements.txt` — Python dependencies.
- `assets/` — local JPG illustrations for the 22 project groups.

Project descriptions are explanatory one-sentence summaries shown on the group tiles; they do not change the saved project titles or student records. For groups whose title is blank, the tile continues to show the existing generic “Project title to be confirmed” title and a general machine-learning description.

## Shared page headings
Every workspace page uses the same rounded, light-gradient title panel, blue chess-piece mark, readable heading, and supporting description. The heading wraps naturally on narrow screens so the title and description are not cropped.

## Run on Windows
1. Extract the ZIP file.
2. Open Command Prompt or PowerShell in the extracted folder.
3. Install dependencies:

   ```bash
   python -m pip install -r requirements.txt
   ```

4. Run the application:

   ```bash
   python -m streamlit run app.py
   ```

Keep `ml_project_evaluation.db`, `student_roster.csv`, and the `assets` folder alongside `app.py` so the application can load its data and local illustrations.

## Project Partners page redesign
The Project Group Details page now uses a larger, responsive project header with the group name, project title, student count and partner initials. Student partner cards have larger readable typography, stronger spacing, aligned card dimensions, clearer marks/progress indicators, and a prominent Evaluate Student action. The database and roster are included unchanged from the source package.


## Login credentials

The portal opens on the login page. Use one of these demo accounts:

- Username: `nehal` · Password: `nehalsheth123`
- Username: `drashti` · Password: `drashtishah123`
- Username: `om` · Password: `ompatel123`

Run the app from this folder so that `app.py`, `ml_project_evaluation.db`, `student_roster.csv`, and `assets/` stay together:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Viva Questions (Group-wise and Phase-wise)

The sidebar now includes **Viva Questions** with two views:
- **Group-wise Viva Questions**: select a group to see its project-specific questions for all seven evaluation phases.
- **Phase-wise Viva Questions**: select an evaluation phase and expand each group to review its five questions for that phase.

The complete Excel question bank is bundled in `ML_Project_Title_Wise_Viva_Questions_22_Groups.xlsx` and can also be downloaded from the Viva Questions page. Keep this workbook in the same folder as `app.py`.

Run with:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```


### Fix for “Missing optional dependency openpyxl”

The app now includes a CSV fallback for the Viva Question Bank, so the Viva Questions page can still load even if `openpyxl` is not installed. The file `ML_Project_Title_Wise_Viva_Questions_22_Groups.csv` must remain beside `app.py`. To enable Excel workbook reading as well, run:

```bash
python -m pip install openpyxl
```

If your computer has multiple Python installations, use `python -m pip` (not just `pip`) with the same Python used to launch Streamlit.


## Viva Questions navigation update

Viva Questions is no longer a separate sidebar navigation item. Open **Student Evaluation**, choose a project group, and use the **Viva Questions · Phase N** link directly beneath the Evaluation Components row. The link opens the selected group's project-specific questions for the active phase. The question bank is read from the bundled CSV, so this feature does not require `openpyxl`.


### Class marksheet exports
The Records & Export page displays each student's marks out of 100 and the half-weighted marks out of 50. It supports Excel (.xlsx) and PDF marksheet downloads, plus the existing CSV and SQLite exports. PDF generation uses ReportLab, installed through `requirements.txt`.


**Excel export dependency:** Install dependencies using `python -m pip install -r requirements.txt`. If Excel export reports that `openpyxl` is missing, run `python -m pip install openpyxl` using the same Python interpreter that launches Streamlit. The app now handles a missing package without crashing; PDF and CSV exports remain available.


## Phase-total evaluation workflow
Faculty enter one editable total per student for each evaluation phase. Individual criteria and their maximum marks remain visible as read-only assessment guidance; criterion-level increment/decrement controls have been removed. Existing criterion marks are migrated to their summed phase total when the evaluation is first loaded, and future marks are saved as phase totals. Class marksheets use the phase totals and calculate the internal score as half of the total out of 100.

## Editable phase-total stepper
The phase-total field now has a red decrement button on the left, an editable numeric marks field in the center, and a green increment button on the right. The buttons respect the phase minimum (0) and maximum marks. This control is available in both individual-student evaluation and the group-wide phase-total entry view.
