# Palomar ACORD Reader & Simple Extractor

## Project Description & Purpose

This Streamlit application extracts structured data from ACORD 129 PDF forms. It parses vehicle and garage location information from uploaded ACORD PDFs and outputs a clean Excel workbook with two sheets: **Vehicles** and **Garages**.

The primary goal is to streamline underwriting for Commerical Auto by automating the extraction of structured data from ACORD forms. This reduces manual entry, minimizes errors, and saves time for underwriters.

## Installation Instructions

This tool is accessible as an internal Palomar webapp at {url} (once we deploy it on Azure). Follow the below instructions for installation on a local environment. 

1. **Clone the repository**

```bash
  git clone <repo_url>
  cd <repo_folder>
```

2. **Set up a virtual environment**

```bash    
python -m venv venv
source venv/bin/activate         # On Windows: venv\Scripts\activate
```

3. **Install dependencies**

```bash
pip install -r requirements.txt
```

---

## Usage Guide

1. **Run the app**

```bash
streamlit run app.py
```

2. **Using the interface**

* Upload an ACORD 129 PDF using the file uploader.
* Wait for the processing to complete.
* Download the resulting Excel file.
* Optionally, preview the parsed vehicle and garage tables via expandable panels.


## Features Overview

* **Vehicle Data Extraction**

  * Parses vehicle number, year, make, model, VIN, address, and cost.

* **Garage Location Parsing**

  * Identifies and extracts structured location details for each garage listed.

* **Excel Export**

  * Creates a downloadable Excel workbook with pre-defined headers and formatting.

* **Streamlit UI**

  * Simple web interface for non-technical users.
  * Expandable panels for previewing parsed data.

* **Stateless**

  * Does not persist uploaded or generated files on disk.


## Notes

* The parsing logic is tuned to common formatting patterns in ACORD 129 forms. Variations may require regex adjustments.
* All processing is done in-memory using `pdfplumber`, `pandas`, and `openpyxl`.


