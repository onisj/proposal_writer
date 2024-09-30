

import os
import re
import docx
import json
import spacy
import shutil
import openai
import tempfile
import streamlit as st
from io import BytesIO
from docx import Document
from PyPDF2 import PdfReader
from dotenv import load_dotenv
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE
from src.constants import GPTModelLight, GPTModel

from src.utils import (
    extract_text_from_pdf, 
    read_word, 
    cleaner, 
    header, 
    writer,
    create_docx_table
)

from src.sections import (
    cover_letter, 
    FIRM_QUALIFICATIONS, 
    TEAM_QUALIFICATIONS, 
    Project_understanding, 
    Technical_approach, 
    Work_plan, 
    COST_PROPOSAL, 
    ref
)



load_dotenv()
OPENAI_API_KEY= os.getenv("OPENAI_API_KEY")

# Load the spaCy model
nlp = spacy.load("en_core_web_sm")

def extract_contact_info(text):
    # Regex patterns for email and phone number
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    phone_pattern = r'(\+?\d{1,3}[-.\s]?)?\(?\d{1,4}?\)?[-.\s]?\d{1,4}[-.\s]?\d{1,9}'

    # Find all emails and phone numbers
    emails = re.findall(email_pattern, text)
    phones = re.findall(phone_pattern, text)

    return emails, phones

def main():
    st.title("Proposal Writer")
    
    # Add a space between the title and the subheader
    st.write("")  
    st.write("")  
    
    # File upload for RFP document
    st.subheader("Upload Documentation")
    uploaded_rfp = st.file_uploader("Files allowed: PDF or DOCX", type=["pdf", "docx"])
    
    purchasing_manager_details = []  
    scope_of_work = ""  

    if uploaded_rfp is not None:
        if uploaded_rfp.type == "application/pdf":
            rfp_text = extract_text_from_pdf(uploaded_rfp)
        elif uploaded_rfp.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            rfp_text = read_word(uploaded_rfp)
        
        # Automatically populate fields based on the extracted text
        if rfp_text:
            # Process the text with spaCy
            doc = nlp(rfp_text)

            # Extract contact info
            emails, phones = extract_contact_info(rfp_text)

            # Define keywords to search for purchasing manager info
            keywords = [
                "purchasing manager", 
                "procurement manager", 
                "procurement coordinator",
                "contact information",
                "department"
            ]
            
            for keyword in keywords:
                start_index = 0  
                while True:
                    start_index = rfp_text.lower().find(keyword, start_index)
                    if start_index == -1:
                        break  
                    
                   
                    context_start = max(0, start_index - 15)  
                    context_end = min(len(rfp_text), start_index + len(keyword) + 200) 
                    
                  
                    purchasing_manager_details.append(rfp_text[context_start:context_end].strip())
                    
                    
                    start_index += len(keyword)

            # Filter out irrelevant information
            purchasing_manager_details = [detail for detail in purchasing_manager_details if "do not contact" not in detail.lower()]
            
            # Store in session state
            st.session_state.purchasing_manager_details = "\n\n".join(purchasing_manager_details)  # Join instances with new lines
    
    # Purchasing Manager Details Section
    if uploaded_rfp is not None:
        st.subheader("Purchasing Manager Details")
        if 'purchasing_manager_details' in st.session_state:
            purchasing_manager_details = st.session_state.purchasing_manager_details
        st.text_area("", value=purchasing_manager_details, height=150)
        
        # Success message for RFP extraction
        if rfp_text:
            st.success("Purchasing Manager extracted successfully.")
        else:
            st.warning("Purchasing Manager not found in the RFP text.")
    
    # Scope of work / service Section
    if uploaded_rfp is not None:
        st.subheader("Scope of Work / Service")
        st.text_area("Extracted Text Below:", value=rfp_text, height=300)

        # Success message for RFP extraction
        if rfp_text:
            st.success("Scope of work / service extracted successfully.")
        else:
            st.warning("Scope of work / service not found in the RFP text.")



    # Article About Disaster Details
    st.subheader("Article About The Disaster")
    article_disaster = st.text_area("Article about the disaster", height=150)



    # Create the 'files' directory if it doesn't exist
    if not os.path.exists("files"):
        os.makedirs("files")

    ## Client Logo Upload
    st.subheader("Client Logo")
    client_logo = st.file_uploader("Upload your logo here", type=["png", "jpg", "jpeg"])
    if client_logo is not None:
        try:
            file_path = os.path.join("files", client_logo.name)
            with open(file_path, "wb") as f:
                f.write(client_logo.getbuffer())
            # relative_path= f"files\{client_logo.name}"
            relative_path = os.path.join("files", client_logo.name)  #this is for both linux and windows
        except Exception:
            st.error("File not uploaded properly.")
    else:
        st.warning("No logo uploaded. Proceeding without a client logo ?")
        relative_path = None  # set default logo path if required #

          
        
    # Yearly Budget Input    
    st.subheader("Yearly Budget")
    yearly_budget = st.text_input("Yearly Budget")
    #uploaded_file = st.file_uploader("Upload a RFP", type="pdf") 


    # Other UI components remain unchanged...  
    if st.button("Generate Proposal"):
        # Generate content
        cl1 = cover_letter(purchasing_manager_details, scope_of_work)
        fq1 = FIRM_QUALIFICATIONS(scope_of_work)
        tq1 = TEAM_QUALIFICATIONS(scope_of_work)
        pu1 = Project_understanding(scope_of_work, article_disaster)
        ta1 = Technical_approach(scope_of_work)
        wp1 = Work_plan(scope_of_work)
        cp1 = COST_PROPOSAL(scope_of_work, yearly_budget)
        r1 = ref(scope_of_work)


        # Create tables
        ta1new = create_docx_table(ta1, 'technical_approach')
        wp1new = create_docx_table(wp1, 'work_plan')
        cp1new = create_docx_table(cp1, 'cost_proposal')

        # Clean content
        cl = cleaner(cl1)
        fq = cleaner(fq1)
        tq = cleaner(tq1)
        pu = cleaner(pu1)
        r = cleaner(r1)

        section5_content = [('text', "Project Understanding:\n" + pu + "\n\nTechnical Approach:")]
        section5_content.extend(ta1new)
        section5_content.append(('text', "\nWork Plan:"))
        section5_content.extend(wp1new)
        
        # Paths
        template_path = "template_proposal.docx"
        output_path = "proposal.docx"

        # Copy template
        shutil.copy(template_path, output_path)
        header(output_path, relative_path)
 
        # Modifications
        modifications = [
            ("COVER LETTER", cl),
            ("SECTION 3: FIRM QUALIFICATIONS, EXPERTISE, AND EXPERIENCE ", fq),
            ("SECTION 4: TEAM QUALIFICATIONS, EXPERTISE, AND EXPERIENCE ", tq),
            ("SECTION 5: TECHNICAL APPROACH AND WORK PLAN", ta1new),
            ("SECTION 5: TECHNICAL APPROACH AND WORK PLAN", wp1new),
            ("SECTION 5: TECHNICAL APPROACH AND WORK PLAN", pu),
            ("SECTION 6: COST PROPOSAL", cp1new),
            ("SECTION 7: REFERENCES ", r),
        ] 

 
        # Apply modifications to the document
        for title, content in modifications:
            print(f"Applying modification for {title}")
            writer(output_path, title, content)

        # Prepare for download
        output = BytesIO()
        op = Document(output_path)
        op.save(output)
        output.seek(0)
        st.download_button(
            label="Download Word File",
            data=output,
            file_name="proposal.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        os.remove(output_path)
        if relative_path is not None:
            os.remove(relative_path)

if __name__ == "__main__":
    main()
