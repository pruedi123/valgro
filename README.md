# Large Growth vs Small Value (LGSV)

This repo now contains a single Streamlit app (`lgsv.py`) that visualizes CAGR differences between the Large Growth and Small Value indexes using the data in `lgsv.xlsx`.

## Setup

1. Create and activate a virtual environment (optional but recommended).
2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Run the app

```bash
streamlit run lgsv.py
```

Streamlit will open a browser tab where you can adjust the moving-average period and select which lines to show. The source data lives in `lgsv.xlsx`, so replace that file if you want to refresh the analysis.
