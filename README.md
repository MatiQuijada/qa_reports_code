# SAF and Banner Reports Comparator

This project is a Streamlit web application designed to compare student report data from two different systems: SAF and Banner. It highlights discrepancies in key fields such as name, email, degree title, major, and thesis advisors.

## Features

- **Upload and compare**: Upload Excel files from SAF and Banner to automatically compare student data.
- **Data normalization**: Handles differences in column names, text formatting, and career naming conventions.
- **Field-by-field comparison**: Checks for matches in Name, Email, Degree Title, Major, Internal Advisor, and External Advisor.
- **Discrepancy highlighting**: Generates an Excel report with mismatches highlighted for easy review.
- **User-friendly interface**: Simple upload and download workflow using Streamlit.

## Requirements

- Python 3.8 or higher
- [pandas](https://pandas.pydata.org/)
- [streamlit](https://streamlit.io/)
- [xlsxwriter](https://xlsxwriter.readthedocs.io/)
- [openpyxl](https://openpyxl.readthedocs.io/) (for reading Excel files)

Install dependencies with:

```sh
pip install -r requirements.txt
```

Example `requirements.txt`:
```
pandas
streamlit
xlsxwriter
openpyxl
```

## How to Use

1. **Run the application:**
   ```sh
   streamlit run src/QA_REPORTS_CODE.py
   ```

2. **Upload files:**
   - Upload the SAF report (Excel file).
   - Upload the Banner report (Excel file).

3. **Review results:**
   - The app will display a table showing the comparison results.
   - Each key field will have a column indicating if the data matches between the two systems.

4. **Download discrepancies:**
   - Click the "Download Discrepancies in Excel" button to get a formatted Excel file.
   - Discrepancies are highlighted for easy identification.

## How It Works

- **Column Mapping:** The app automatically renames columns from both sources to a standard format.
- **Normalization:** Text fields are lowercased, accents are removed, and whitespace is trimmed.
- **Career Equivalency:** Different names for the same major are mapped to a standard value.
- **Comparison:** Each student (matched by RUT) is compared field by field. Names are compared word-by-word for similarity.
- **Output:** The result is shown in the app and can be downloaded as an Excel file with highlighted mismatches.

## File Structure

```
src/
  QA_REPORTS_CODE.py   # Main Streamlit app
requirements.txt       # Python dependencies
```

## Author

Developed by [Your Name].

---

*For questions or suggestions, please open an issue or contact the author.*