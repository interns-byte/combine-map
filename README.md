# AUiX Network - v7

This version connects the two AUiX stakeholder maps inside one Streamlit application.

## Views

- **By Organization Type**: AUiX in the center, organization categories on the inner ring, and all organizations on the fixed outer ring. Click a category to focus its organizations.
- **By Expertise**: AUiX in the center, expertise areas on the inner ring, and the same organizations in the same fixed outer-ring positions. Select up to three expertise areas at once.

## Shared behavior

- Organization search works across both views.
- Up to four organization profiles remain pinned when switching views.
- The same Excel workbook powers both views.
- On mobile, the network is full width and pinned profiles move below the map.

## Run locally

```bash
python3 -m pip install -r requirements.txt
python3 -m streamlit run app.py
```

## Streamlit Community Cloud

Replace the existing `app.py` in your GitHub repository with this version and commit the change. The existing workbook can stay where it is; the app checks both `data/AUiX_Expertise_Map_Data.xlsx` and the repository root.

## v9.4 header branding
- Adds centered DRADIS and AUiX logos above the page title.
- Adds the note: “Best viewed on desktop or tablet.”
- Logo PNGs are included in the `assets/` folder and must be committed with `app.py`.

