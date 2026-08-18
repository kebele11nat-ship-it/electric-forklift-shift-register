# Electric Forklift Shift Register

A browser-based OC shift-entry tool for recording the status of the electric forklift fleet at the beginning of every shift.

## Main purpose
The OC enters each forklift's:
- Status
- Battery charge %
- Operator assigned
- Operation / location (L1, L2, RM, Inside, etc.)

The app immediately summarizes:
- Total forklifts
- Active forklifts
- Charging forklifts
- Average battery charge
- Operators assigned
- Operators missing
- Forklifts with low battery

## Records and Excel
Each saved shift is stored in the browser's local storage and can be exported to an Excel workbook containing the full shift register and summary.

## GitHub Pages
The direct browser app is `index.html`, so it can run from GitHub Pages without Python or Streamlit.

## Next upgrade
For a shared permanent company-wide register across multiple devices, connect the form to Supabase (free tier) and automatically synchronize every submitted shift to the central database.
