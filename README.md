# iMocha PPT Auto Formatting Tool

A simple Flask web app that accepts `.ppt` and `.pptx` files and converts them into a branded iMocha presentation.

## Features

- Upload a PowerPoint file and receive a reformatted `.pptx` file.
- Applies iMocha brand colors, fonts, logo placement, and consistent slide structure.
- Uses sample design assets from `Sample_Ppt.pptx` and `iMocha_logo.png`.

## Setup

1. Create and activate a Python virtual environment.
2. Install dependencies:
   ```bash
   python -m pip install -r requirements.txt
   ```
3. Run the app:
   ```bash
   python app.py
   ```
4. Open `http://127.0.0.1:5000` in your browser.

## Usage

- Upload a `.ppt` or `.pptx` file using the web UI.
- Download the generated `iMocha_Formatted_Presentation.pptx`.

## Notes

- The tool extracts text from the uploaded presentation and restructures it onto new slides.
- The first slide uses the iMocha logo and the sample extracted image when available.
