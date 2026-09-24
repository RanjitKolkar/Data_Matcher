# Excel Intelligence & Matcher

A deployable Streamlit application for loading multiple Excel workbooks and comparing selected columns.

## Features

- Multiple Excel workbook upload
- Multiple sheets per workbook
- Workbook and sheet statistics
- Row / column / blank / distinct-value statistics
- Dynamic column mapping
- Common entries
- Distinct entries
- Missing-by-source analysis
- Fuzzy/similar matching
- Adjustable similarity threshold
- Token Sort, Token Set and WRatio matching
- Search/filter results
- Similarity groups
- Source row/location tracking
- Highlighted results
- Complete Excel report export
- CSV export

## Local installation

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload `app.py` and `requirements.txt`.
3. Open Streamlit Community Cloud.
4. Select the GitHub repository.
5. Set the main file to `app.py`.
6. Deploy.

No database or external API is required.
