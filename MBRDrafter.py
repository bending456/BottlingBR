## Opening of Python Packages to run the program ###
import streamlit as st
import os
import tempfile
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import Inches

st.header("Master Batch Record Drafter [Under Construction]")
document = Document()

# Create a new style for each indent level
for i in range(5):  # Adjust range for as many levels as you need
  try:
      style = document.styles.add_style(f'List Bullet {i}', document.styles['Normal'].type)
  except:
      style = document.styles[f'List Bullet {i}']
  style.paragraph_format.left_indent = Pt(18 * i)  # 36 points = 0.5 inches
  style.paragraph_format.first_line_indent = Pt(-9)  # 18 points = 0.25 inches
  style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
  style.font.size = Pt(11)
# Header of Document
title = document.add_paragraph()
titleText = st.text_input("Write the title of your document in here ... ")
outputfileName = st.text_input("Write the name of output (docx file) name in here ...")
run = title.add_run(titleText)
run.bold = True
run.font.size = Pt(18)

st.markdown("--------------------------")
bundling = st.checkbox("Bundling?")
if bundling:
  st.caption("Bundling is selected")
  p = document.add_paragraph(style=document.styles['List Bullet 0'])
  p.paragraph_format.line_spacing = Pt(12)  # Set line spacing to 24 points
  # Main Process Name
  p.add_run('Bundling')
  bparentstep1 = st.checkbox('Step: Parent Bundling Step 1')

  if bparentstep1:
     p = document.add_paragraph(style=document.styles['List Bullet 1'])
     p.add_run('Bundling Parent Step 1')
     st.caption('- Choose specific bundling steps')
     
     bchildstep1_1 = st.checkbox('Sub step: child bundling step 1-1')
     if bchildstep1_1:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('Bundling Child Step 1-1')

     bchildstep1_2 = st.checkbox('Sub step: child bundling step 1-2')
     if bchildstep1_2:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('Bundling Child Step 1-2')
  
  st.markdown("--------------------------")
  bparentstep2 = st.checkbox('Step: Parent Bundling Step 2')
  if bparentstep2:
     p = document.add_paragraph(style=document.styles['List Bullet 1'])
     p.add_run('Bundling Parent Step 2')
     st.caption('- Choose specific bundling steps')
     
     bchildstep2_1 = st.checkbox('Sub step: child bundling step 2-1')
     if bchildstep2_1:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('Bundling Child Step 2-1')

     bchildstep2_2 = st.checkbox('Sub step: child bundling step 2-2')
     if bchildstep2_2:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('Bundling Child Step 2-2')
  
#-------------------------------------------------------------------------------
st.markdown("--------------------------")
cartoning = st.checkbox("Cartoning?")
if cartoning:
  st.caption("cartoning is selected")
  p = document.add_paragraph(style=document.styles['List Bullet 0'])
  p.paragraph_format.line_spacing = Pt(12)  # Set line spacing to 24 points
  # Main Process Name
  p.add_run('cartoning')

  cparentstep1 = st.checkbox('Step: Parent cartoning Step 1')  
  if cparentstep1:
     p = document.add_paragraph(style=document.styles['List Bullet 1'])
     p.add_run('cartoning Parent Step 1')
     st.caption('- Choose specific bundling steps')
     
     cchildstep1_1 = st.checkbox('Sub step: child cartoning step 1-1')
     if cchildstep1_1:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('cartoning Child Step 1-1')

     cchildstep1_2 = st.checkbox('Sub step: child cartoning step 1-2')
     if cchildstep1_2:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('cartoning Child Step 1-2')
  
  st.markdown("--------------------------")
  cparentstep2 = st.checkbox('Step: Parent cartoning Step 2')
  if cparentstep2:
     p = document.add_paragraph(style=document.styles['List Bullet 1'])
     p.add_run('cartoning Parent Step 2')
     st.caption('- Choose specific bundling steps')

     cchildstep2_1 = st.checkbox('Sub step: child cartoning step 2-1')
     if cchildstep2_1:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('cartoning Child Step 2-1')

     cchildstep2_2 = st.checkbox('Sub step: child cartoning step 2-2')
     if cchildstep2_2:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('cartoning Child Step 2-2')

#-------------------------------------------------------------------------------
st.markdown("--------------------------")
additional = st.checkbox("Additional?")
if additional:
  st.caption("additional is selected")
  p = document.add_paragraph(style=document.styles['List Bullet 0'])
  p.paragraph_format.line_spacing = Pt(12)  # Set line spacing to 24 points
  # Main Process Name
  p.add_run('additional')

  st.markdown("--------------------------")
  aparentstep1 = st.checkbox('Step: Parent additional Step 1')
  if aparentstep1:
     p = document.add_paragraph(style=document.styles['List Bullet 1'])
     p.add_run('additional Parent Step 1')
     st.caption('- Choose specific bundling steps')
     
     achildstep1_1 = st.checkbox('Sub step: child additional step 1-1')
     if achildstep1_1:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('additional Child Step 1-1')

     achildstep1_2 = st.checkbox('Sub step: child additional step 1-2')
     if achildstep1_2:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('additional Child Step 1-2')
  
  st.markdown("--------------------------")
  aparentstep2 = st.checkbox('Step: Parent additional Step 2')
  if aparentstep2:
     p = document.add_paragraph(style=document.styles['List Bullet 1'])
     p.add_run('additional Parent Step 2')
     st.caption('- Choose specific bundling steps')

     achildstep2_1 = st.checkbox('Sub step: child additional step 2-1')
     if achildstep2_1:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('additional Child Step 2-1')

     achildstep2_2 = st.checkbox('Sub step: child additional step 2-2')
     if achildstep2_2:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('additional Child Step 2-2')
  # Save the document
#document.save(outputfileName+'.docx')
with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
    document.save(tmp.name)
    tmp.seek(0)

    # Create a button to download the docx file
    st.download_button(
        label="Download .docx file",
        data=tmp.read(),
        file_name=outputfileName+".docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

# Remove temporary file
os.unlink(tmp.name)
