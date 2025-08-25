import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader
import pdfplumber
from io import BytesIO
from PIL import Image
import re


class AcordParser:
    """
    Extracts AcroForm fields if available.
    If missing, falls back to pdfplumber text parsing (page 21).
    """

    FIELD_MAP = [
        "Agency Name",
        "Agency Phone Number",
        "Physical Address",
        "City",
        "Zip Code",
        "State",
        "Mail Address If different from physical address",
        "City_2",
        "Zip Code_2",
        "State_2",
        "Operations Contact Name",
        "Operations Contact Email automated policy related emails will be sent to this email",
        "Accounting Contact Name",
        "Accounting Contact Email used for commission statement delivery",
        "Agency License Number",
        "Agency License Number_2",
        "Agency National Producer Number NPN",
        "Main Producer Name",
        "Main Producer Email",
        "Main Producer National Producer Number NPN",
    ]

    @staticmethod
    def parse_flat_text(text: str, source_file: str) -> dict:
        """Parse pdfplumber text block into the same schema as AcroForm fields."""
        data = {}

        def extract(pattern, text, group=1):
            m = re.search(pattern, text, re.IGNORECASE)
            return m.group(group).strip() if m else ""

        # Agency
        data["Agency Name"] = extract(r"Agency Name:\s*(.+)", text)
        data["Agency Phone Number"] = extract(r"Agency Phone Number:\s*([\d\.\-\(\)\s]+)", text)
        data["Physical Address"] = extract(r"Physical Address:\s*(.+)", text)
        data["City (Physical)"] = extract(r"City:\s*([A-Za-z\s]+)\s*Zip Code:", text)
        data["Zip Code (Physical)"] = extract(r"Zip Code:\s*([\d\-]+)\s*State:", text)
        data["State (Physical)"] = extract(r"State:\s*([A-Z]{2})", text)

        # Mailing
        data["Mail Address"] = extract(r"Mail Address.*?:\s*(.+)", text)
        data["City (Mailing)"] = extract(r"Mail Address.*?\nCity:\s*([A-Za-z\s]+)\s*Zip Code:", text)
        data["Zip Code (Mailing)"] = extract(r"Mail Address.*?\nCity:.*?Zip Code:\s*([\d\-]+)\s*State:", text)
        data["State (Mailing)"] = extract(r"Mail Address.*?\nCity:.*?State:\s*([A-Z]{2})", text)


        # Contacts
        data["Operations Contact Name"] = extract(r"Operations Contact Name:\s*(.+)", text)
        data["Operations Contact Email"] = extract(
            r"Operations Contact Email.*?:\s*([\w\.-]+@[\w\.-]+)", text
        )
        data["Accounting Contact Name"] = extract(r"Accounting Contact Name:\s*(.+)", text)
        data["Accounting Contact Email"] = extract(
            r"Accounting Contact Email.*?:\s*([\w\.-]+@[\w\.-]+)", text
        )

        # Licensing
        data["Agency License State"] = extract(r"Agency License State:\s*([A-Z]{2})", text)
        data["Agency License Number"] = extract(r"Agency License Number:\s*([\w\d]+)", text)
        data["Agency National Producer Number (NPN)"] = extract(
            r"Agency National Producer Number.*?:\s*([\d]+)", text
        )

        # Main Producer
        data["Main Producer Name"] = extract(r"Main Producer Name:\s*(.+?)\s+Main Producer Email:", text)
        data["Main Producer Email"] = extract(r"Main Producer Email:\s*([\w\.-]+@[\w\.-]+)", text)
        data["Main Producer NPN"] = extract(
            r"Main Producer National Producer Number.*?:\s*([\d]+)", text
        )

        data["Source File"] = source_file
        return data

    @staticmethod
    def extract(pdf_file, page_number: int = 21) -> dict:
        reader = PdfReader(pdf_file)
        fields = reader.get_fields() or {}

        # Check if any of our desired fields exist
        found = any(f in fields for f in AcordParser.FIELD_MAP)

        if found:
            values = {}
            for f in AcordParser.FIELD_MAP:
                v = fields.get(f, {})
                val = ""
                if v:
                    val = v.get("/V") or v.get("V") or ""
                values[f] = str(val) if val is not None else ""

            values["Source File"] = pdf_file.name
            return values

        else:
            # Fallback to pdfplumber text parsing
            with pdfplumber.open(pdf_file) as pdf:
                if page_number - 1 < len(pdf.pages):
                    page = pdf.pages[page_number - 1]
                    text = page.extract_text() or ""
                else:
                    text = ""
            return AcordParser.parse_flat_text(text, pdf_file.name)

    @staticmethod
    def generate_excel(all_forms: pd.DataFrame) -> BytesIO:
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            all_forms.to_excel(writer, sheet_name="Results", index=False)
        output.seek(0)
        return output


def main():
    st.set_page_config(page_title="New Producer Parser", layout="centered")

    try:
        logo = Image.open("PLMR_BIG.png")
        st.image(logo, use_container_width=True)
    except Exception:
        pass

    st.markdown("<h2 style='text-align: center;'>New Producer Parser</h2>", unsafe_allow_html=True)
    st.markdown("---")

    uploaded_files = st.file_uploader(
        "Upload one or more PDFs",
        type=["pdf"],
        accept_multiple_files=True
    )

    if uploaded_files:
        all_data = []
        with st.spinner("Extracting data..."):
            for file in uploaded_files:
                parsed = AcordParser.extract(file, page_number=21)
                all_data.append(parsed)

        df = pd.DataFrame(all_data)
        excel_file = AcordParser.generate_excel(df)

        st.success("Extraction complete.")

        st.download_button(
            label="Download Excel File",
            data=excel_file,
            file_name="all_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.dataframe(df)


if __name__ == "__main__":
    main()
