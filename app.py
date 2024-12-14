from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import streamlit as st
from datetime import datetime

def replace_placeholders_sat(doc, placeholders):
    """Replace placeholders in a Word document, maintaining original formatting."""
    def replace_in_paragraph(paragraph, key, value):
        for run in paragraph.runs:
            if key in run.text:
                run.text = run.text.replace(key, value)
                run.font.name = paragraph.style.font.name
                run.font.size = paragraph.style.font.size

    def replace_in_cell(cell, placeholders):
        for para in cell.paragraphs:
            for key, value in placeholders.items():
                replace_in_paragraph(para, key, value)

    for para in doc.paragraphs:
        for key, value in placeholders.items():
            replace_in_paragraph(para, key, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_in_cell(cell, placeholders)

    return doc

def replace_placeholders(doc, placeholders):
    """Replace placeholders in a Word document, including paragraphs and tables."""
    
    def replace_in_paragraph(paragraph, key, value):
        """Replace placeholders in a single paragraph, handling split runs."""
        full_text = "".join(run.text for run in paragraph.runs)
        if key in full_text:
            full_text = full_text.replace(key, value)
            for run in paragraph.runs:
                run.text = ""  # Clear all runs
            paragraph.runs[0].text = full_text  # Add the replaced text back

    def replace_in_cell(cell, placeholders):
        """Replace placeholders inside a table cell."""
        for para in cell.paragraphs:
            for key, value in placeholders.items():
                replace_in_paragraph(para, key, value)

    # Replace placeholders in all paragraphs
    for para in doc.paragraphs:
        for key, value in placeholders.items():
            replace_in_paragraph(para, key, value)

    # Replace placeholders in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_in_cell(cell, placeholders)

    return doc
def apply_image_placeholder(doc, placeholder_key, image_file):
    """Replace a placeholder with an image in the Word document."""
    try:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if placeholder_key in para.text:
                            para.text = ""
                            run = para.add_run()
                            run.add_picture(image_file, width=Inches(1.5), height=Inches(0.75))
                            return doc
        for para in doc.paragraphs:
            if placeholder_key in para.text:
                para.text = ""
                run = para.add_run()
                run.add_picture(image_file, width=Inches(1.5), height=Inches(0.75))
                return doc
        raise ValueError(f"Placeholder '{placeholder_key}' not found in the document.")
    except Exception as e:
        raise Exception(f"Error inserting image: {e}")

st.title("Dynamic Document Generator")

template_option = st.selectbox("Select Template", ["SAT Template", "Service Agreement"])

if template_option == "SAT Template":
    template_path = "SAMPLE VAT registration and VAT filling -SME package.docx"
    doc = Document(template_path)

    agreement_date = st.date_input("Date of Agreement", datetime.today())
    reference_number = st.text_input("Reference Number")
    atten = st.text_input("Atten")
    email = st.text_input("Email")
    client_name = st.text_input("Client Name")
    commercial_registration_number = st.text_input("Commercial Registration Number")
    service_provider_name = st.text_input("Service Provider Name")
    service_provider_cr = st.text_input("Service Provider CR Number")
    company_name = st.text_input("Company Name")
    vat_registration_fee = st.text_input("VAT Registration Fee")
    consultancy_fee = st.text_input("Consultancy Fee")
    authorized_person_name = st.text_input("Authorized Person Name")
    signature_image = st.file_uploader("Upload Signature Image", type=["png", "jpg", "jpeg"])

    placeholders = {
        "<<Date>>": agreement_date.strftime("%d-%m-%Y"),
        "<<Reference Number>>": reference_number,
        "<<Atten>>": atten,
        "<<Email>>": email,
        "<<Client Name>>": client_name,
        "<<Commercial Registration Number>>": commercial_registration_number,
        "<<Service Provider Name>>": service_provider_name,
        "<<Service Provider CR Number>>": service_provider_cr,
        "<<Company Name>>": company_name,
        "<<VAT Registration Fee>>": vat_registration_fee,
        "<<Consultancy Fee>>": consultancy_fee,
        "<<Authorized Person Name>>": authorized_person_name,
    }

    if st.button("Generate SAT Document"):
        try:
            doc = replace_placeholders_sat(doc, placeholders)
            if signature_image:
                doc = apply_image_placeholder(doc, "<<Signature Image>>", signature_image)

            formatted_date = agreement_date.strftime("%d %b %Y")
            output_path = f"SAT - {client_name} {formatted_date}.docx"
            doc.save(output_path)

            st.success("Document generated successfully!")
            with open(output_path, "rb") as file:
                st.download_button("Download Document", data=file, file_name=output_path)

        except Exception as e:
            st.error(f"An error occurred: {e}")

elif template_option == "Service Agreement":
    template_path = "SAMPLE Service Agreement -Company formation -Bahrain - Filled.docx"
    doc = Document(template_path)

    agreement_date = st.date_input("Date of Agreement", datetime.today())
    reference_number = st.text_input("Reference Number")
    client_name = st.text_input("Client Name")
    bahraini_ownership = st.number_input("Bahraini Ownership (%)", min_value=0, max_value=100, step=1)
    gcc_ownership = st.number_input("GCC Nationals Ownership (%)", min_value=0, max_value=100, step=1)
    american_ownership = st.number_input("American Nationals Ownership (%)", min_value=0, max_value=100, step=1)
    foreign_ownership = st.number_input("Foreign Ownership (%)", min_value=0, max_value=100, step=1)

    business_activity_1_isic = st.text_input("Business Activity ISIC4 Code (1st)")
    business_activity_1_name = st.text_input("Business Activity Name (1st)")
    business_activity_1_desc = st.text_area("Business Activity Description (1st)")
    business_activity_2_isic = st.text_input("Business Activity ISIC4 Code (2nd)")
    business_activity_2_name = st.text_input("Business Activity Name (2nd)")
    business_activity_2_desc = st.text_area("Business Activity Description (2nd)")

    costs = {
        "Company Formation Cost": st.number_input("Company Formation Cost", min_value=0.0, step=0.01),
        "Desk-Space Office Rental Cost": st.number_input("Desk-Space Office Rental Cost", min_value=0.0, step=0.01),
        "Businessman Visa Cost": st.number_input("Businessman Visa Cost", min_value=0.0, step=0.01),
        "Miscellaneous/Admin Charges": st.number_input("Miscellaneous/Admin Charges", min_value=0.0, step=0.01),
        "Power of Attorney Cost": st.number_input("Power of Attorney Cost", min_value=0.0, step=0.01),
        "Estimation Charges (Per Head)": st.number_input("Estimation Charges (Per Head)", min_value=0.0, step=0.01),
        "Labour Authority Registration Cost": st.number_input("Labour Authority Registration Cost", min_value=0.0, step=0.01),
        "Social Insurance Registration Cost": st.number_input("Social Insurance Registration Cost", min_value=0.0, step=0.01),
        "Free Advice/Guidance Cost": st.number_input("Free Advice/Guidance Cost", min_value=0.0, step=0.01),
    }

    total_cost = sum(costs.values())

    signatory_name = st.text_input("Signatory Name")
    passport_number = st.text_input("Passport Number")
    signature_image = st.file_uploader("Upload Signature Image", type=["png", "jpg", "jpeg"])

    placeholders = {
        "<< Date >>": agreement_date.strftime("%d-%m-%Y"),
        "<< Reference Number >>": reference_number,
        "<< Client Name >>": client_name,
        "<< Bahraini Ownership >>": f"{bahraini_ownership}%",
        "<< GCC Nationals Ownership >>": f"{gcc_ownership}%",
        "<< American Nationals Ownership >>": f"{american_ownership}%",
        "<< Foreign Ownership >>": f"{foreign_ownership}%",
        "<<Text1>>": business_activity_1_isic,
        "<<Text2>>": business_activity_1_name,
        "<<Text3>>": business_activity_1_desc,
        "<<Text4>>": business_activity_2_isic,
        "<<Text5>>": business_activity_2_name,
        "<<Text6>>": business_activity_2_desc,
        "<< Company Formation Cost >>": f"{costs['Company Formation Cost']:.2f}",
        "<< Desk-Space Office Rental Cost >>": f"{costs['Desk-Space Office Rental Cost']:.2f}",
        "<< Businessman Visa Cost >>": f"{costs['Businessman Visa Cost']:.2f}",
        "<< Miscellaneous/Admin Charges >>": f"{costs['Miscellaneous/Admin Charges']:.2f}",
        "<< Power of Attorney Cost >>": f"{costs['Power of Attorney Cost']:.2f}",
        "<< Estimation Charges (Per Head) >>": f"{costs['Estimation Charges (Per Head)']:.2f}",
        "<< Labour Authority Registration Cost >>": f"{costs['Labour Authority Registration Cost']:.2f}",
        "<< Social Insurance Registration Cost >>": f"{costs['Social Insurance Registration Cost']:.2f}",
        "<< Free Advice/Guidance Cost >>": f"{costs['Free Advice/Guidance Cost']:.2f}",
        "<< Total Cost >>": f"{total_cost:.2f}",
        "<< Signatory Name >>": signatory_name,
        "<< Passport Number >>": passport_number,
    }

    if st.button("Generate Service Agreement Document"):
        try:
            doc = replace_placeholders(doc, placeholders)
            if signature_image:
                doc = apply_image_placeholder(doc, "<< Signatory Image >>", signature_image)

            formatted_date = agreement_date.strftime("%d %b %Y")
            output_path = f"Service Agreement - {client_name} {formatted_date}.docx"
            doc.save(output_path)

            st.success("Document generated successfully!")
            with open(output_path, "rb") as file:
                st.download_button("Download Document", data=file, file_name=output_path)

        except Exception as e:
            st.error(f"An error occurred: {e}")
