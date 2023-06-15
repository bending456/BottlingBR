## Opening of Python Packages to run the program ###
import streamlit as st
import numpy as np
import os
import tempfile
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import Inches
from docx.shared import RGBColor
from docx.oxml import OxmlElement
from datetime import date
import string


##########################################
#----------------------------------------#
#----------------------------------------#
##########################################

st.header("Master Batch Record Drafter")
st.caption("[Not Ready]")

##########################################
#----------------------------------------#
#----------------------------------------#
##########################################

st.markdown("### ***README***")
with st.expander("User Guide",expanded=True):
#   st.caption("Step 0: Check the box to prevent auto reset")
#   st.caption("Step 1: Type Document Name in Sidebar ")
#   st.caption("Step 2: Type Output Name in Sidebar ")
#   st.caption("Step 3: Type Batch Number in Sidebar ")
#   st.caption("Step 4: Type Your Name in Sidebar ")
#   st.caption("Step 5: Select Packaging Process(s) - [List of Processes] ")
#   st.caption("Step 6: Select Specific Steps - [Process Control Panel] ")
   st.caption("Coming soon...")

if 'writing draft' not in st.session_state:
   st.session_state['writing draft']=False

document = Document('UBR.docx')

#---------- Sidebar Setup
stateholder = st.sidebar.checkbox("Step 0: Check this box to prevent unwanted rerun")
if stateholder:
   st.session_state['writing draft']=True


##########################################
#----------------------------------------#
#----------------------------------------#
##########################################

##----------- docx file generator setup 
####------- Create a new style for each indent level
for i in range(5):  # Adjust range for as many levels as you need
  try:
      style = document.styles.add_style(f'List Bullet {i}', document.styles['Normal'].type)
  except:
      style = document.styles[f'List Bullet {i}']
  style.paragraph_format.left_indent = Pt(18 * i)  # 36 points = 0.5 inches
  style.paragraph_format.first_line_indent = Pt(0)  # 18 points = 0.25 inches
  style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
  style.font.size = Pt(11)

def set_col_widths(table):
    widths = (Inches(0.47), Inches(4.69), Inches(0.97), Inches(0.67), Inches(0.67))
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

style = document.styles['Normal']
font = style.font
font.name = 'Times New Roman'
font.size = Pt(11)

sections = document.sections
for section in sections:
    section.top_margin = Inches(0.0)
    section.bottom_margin = Inches(0.0)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)


def remove_table_spacing(doc):
    # Iterate through each paragraph in the document
    for paragraph in doc.paragraphs:
        # Check if the paragraph contains a table
        if paragraph._p.xml.startswith('<w:tbl'):
            # Access the paragraph's paragraph format
            paragraph_format = paragraph.paragraph_format
            # Set the spacing before and after the table to zero
            paragraph_format.space_before = Pt(0)
            paragraph_format.space_after = Pt(0)


## Access the tables in word file 

## Iterate over each table in the Word Doc.
'''
Structure of the template
Table 1: Section1 - Document Approval and Review
Table 2: Section2 - General Information - Table of Content:
Table 3: Section3 - Reference Information - Referenced Documents 
Table 4: Section4 - Primary Packaging Equipment List 
Table 5: Section5 - Primary Packaging Operations #<----- This is where we need to add input via Python

'''

tablecounter1 = 0 #<---- This will count a number of tables being processed.
st.markdown("### ***Setting up Primary Packaging Operations***")
with st.expander("Primary Packaging Operations"):
    st.caption("Fill Count Input")
    fillcount = st.text_input("Fill count per bottle")
    fillcountref = st.text_input("Fill count per bottle reference (ex. PSIS-Sec X)")
    st.divider()
    st.caption("Total Bottle Required Input")
    totalbottle = st.text_input("Total Bottles Required")
    totalbottleref = st.text_input("Total Bottles Required reference (ex. PSIS-Sec X)")
    st.divider()
    st.caption("Verify the status of Wipotec Scale")
    verWipotecref = st.text_input("Verification Reference (ex. PSIS-Sec X)")

for table in document.tables:
    tablecounter1 += 1
    if tablecounter1 == 4:
        cell1 = table.rows.cells(14,3)
        cell2 = table.rows.cells(15,3)
        cell3 = table.rows.cells(16,4)
        cell1.text = 'Fill Count per\n Bottle\n'+fillcount+'\n'+fillcountref
        cell2.text = 'Total Bottles\n Required\n'+totalbottle+'\n'+totalbottleref
        cell3.text = verWipotecref
        
        paragraph = cell1.paragraphs[0]
        run = paragraph.runs
        for run in paragraph.runs:
            run.font.bold = True 
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cell1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        paragraph = cell2.paragraphs[0]
        run = paragraph.runs
        for run in paragraph.runs:
            run.font.bold = True 
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cell2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        paragraph = cell3.paragraphs[0]
        run = paragraph.runs
        for run in paragraph.runs:
            run.font.bold = True 
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cell3.vertical_alignment = WD_ALIGN_VERTICAL.CENTER



##########################################
#----------------------------------------#
#----------------------------------------#
##########################################
remove_table_spacing(document)
# Save the document
st.sidebar.header("**Step 7: Download Ready**")
if st.sidebar.checkbox("Check this box if the draft is ready"):
   with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
       document.save(tmp.name)
       tmp.seek(0)

       # Create a button to download the docx file
       st.sidebar.download_button(
           label="Download .docx file",
           data=tmp.read(),
           file_name="UBRDraft.docx",  ##<----- Make sure you change this to customizable command
           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
       )

   # Remove temporary file
   os.unlink(tmp.name)
