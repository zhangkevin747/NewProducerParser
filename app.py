import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader
from io import BytesIO
from PIL import Image


class AcordParser:
    """
    Extracts form fields from onboarding PDFs
    and outputs them in the same order as the questionnaire form.
    """

    @staticmethod
    def extract_form_fields(pdf_file, page_number: int = None) -> dict:
        reader = PdfReader(pdf_file)
        fields = reader.get_fields() or {}

        # Optionally filter to one page (1-based index for convenience)
        if page_number is not None and 0 <= page_number - 1 < len(reader.pages):
            page = reader.pages[page_number - 1]
            page_annots = page.get("/Annots") or []
            page_field_names = set()

            for annot_ref in page_annots:
                annot = annot_ref.get_object()
                if "/T" in annot:
                    page_field_names.add(annot["/T"])

            # Reduce to only fields from this page
            fields = {k: v for k, v in fields.items() if k in page_field_names}

        # Flatten into dict {field_name: value}
        values = {}
        for name, f in fields.items():
            val = f.get("/V") or f.get("V") or ""
            values[name] = str(val) if val is not None else ""

        # Map into your schema (same as before)
        ordered_data = {
            "Agency Name": values.get("Agency Name", ""),
            "Agency Phone Number": values.get("Agency Phone Number", ""),
            "Physical Address": values.get("Physical Address", ""),
            "City (Physical)": values.get("City", ""),
            "Zip Code (Physical)": values.get("Zip Code", ""),
            "State (Physical)": values.get("State", ""),
            "Mail Address": values.get("Mail Address If different from physical address", ""),
            "City (Mailing)": values.get("City_2", ""),
            "Zip Code (Mailing)": values.get("Zip Code_2", ""),
            "State (Mailing)": values.get("State_2", ""),
            "Operations Contact Name": values.get("Operations Contact Name", ""),
            "Operations Contact Email": values.get("Operations Contact Email automated policy related emails will be sent to this email", ""),
            "Accounting Contact Name": values.get("Accounting Contact Name", ""),
            "Accounting Contact Email": values.get("Accounting Contact Email used for commission statement delivery", ""),
            "Agency License Number": values.get("Agency License Number", ""),
            "Agency License Number (alt)": values.get("Agency License Number_2", ""),
            "Agency National Producer Number (NPN)": values.get("Agency National Producer Number NPN", ""),
            "Main Producer Name": values.get("Main Producer Name", ""),
            "Main Producer Email": values.get("Main Producer Email", ""),
            "Main Producer NPN": values.get("Main Producer National Producer Number NPN", "")
        }

        return ordered_data


    @staticmethod
    def generate_excel(all_forms: pd.DataFrame) -> BytesIO:
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            all_forms.to_excel(writer, sheet_name="FormFields", index=False)
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

    uploaded_files = st.file_uploader("Upload one or more Fillable PDFs", type=["pdf"], accept_multiple_files=True)

    if uploaded_files:
        all_data = []
        with st.spinner("Extracting form fields..."):
            for file in uploaded_files:
                parsed = AcordParser.extract_form_fields(file, page_number=21)
                parsed["Source File"] = file.name  # new role/column to track file of origin
                all_data.append(parsed)

        df = pd.DataFrame(all_data)
        excel_file = AcordParser.generate_excel(df)

        st.success("Extraction complete.")

        st.download_button(
            label="Download Excel File",
            data=excel_file,
            file_name="all_formfields.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.dataframe(df)


if __name__ == "__main__":
    main()
