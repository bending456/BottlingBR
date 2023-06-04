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
from docx.shared import RGBColor

st.header("Master Batch Record Drafter [Under Construction]")
if 'writing draft' not in st.session_state:
   st.session_state['writing draft']=False

document = Document()

#---------- Sidebar Setup
if st.sidebar.checkbox("Check this box to prevent unwanted rerun"):
   st.session_state['writing draft']=True

##----------- docx file generator setup 
####------- Create a new style for each indent level
for i in range(5):  # Adjust range for as many levels as you need
  try:
      style = document.styles.add_style(f'List Bullet {i}', document.styles['Normal'].type)
  except:
      style = document.styles[f'List Bullet {i}']
  style.paragraph_format.left_indent = Pt(18 * i)  # 36 points = 0.5 inches
  style.paragraph_format.first_line_indent = Pt(-9)  # 18 points = 0.25 inches
  style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
  style.font.size = Pt(11)

###------ Title of Document
title = document.add_paragraph()
titleText = st.sidebar.text_input("Write the title of your document in here ... ")
outputfileName = st.sidebar.text_input("Write the name of output (docx file) name in here ...")
run = title.add_run(titleText)
run.bold = True
run.font.size = Pt(14)

###------ Header of Document
header = document.sections[0].header
paragraph = header.paragraphs[0]
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
batchnumber = st.sidebar.text_input("Write the batch number in here ...")
run = paragraph.add_run(batchnumber)
run.bold = True
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(255, 0, 0)

###------ Footer of Document
footer = document.sections[0].footer
paragraph = footer.paragraphs[0]
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
run = paragraph.add_run("Confidential")
run.bold = True
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(0, 0, 255)

##---- Selecting Processes
st.sidebar.header("**List of Processes**")
###----------- List of Processes ---------------------------
bundling = st.sidebar.checkbox("Bundling?",value=True)
cartoning = st.sidebar.checkbox("Cartoning?",value=True)
additional = st.sidebar.checkbox("Additional?",value=True)



if bundling:
  st.markdown("--------------------------")
  st.markdown('<p style="font-size: 20px;">Bundling is selected</p>', unsafe_allow_html=True)
  st.markdown("--------------------------")
  p = document.add_paragraph(style=document.styles['List Bullet 0'])
  p.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 24 points
  # Main Process Name
  p.add_run('Bundling')
  bparentstep1 = st.checkbox('Step: Parent Bundling Step 1',value=True)
  bparentstep1warning = st.checkbox('Any warning regarding bundling step 1?')
  if bparentstep1:
     p = document.add_paragraph(style=document.styles['List Bullet 1'])
     p.add_run('Bundling Parent Step 1')
     st.caption('- Choose specific bundling steps')
     
     bchildstep1_1 = st.checkbox('Sub step: child bundling step 1-1',value=True)
     if bchildstep1_1:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('Bundling Child Step 1-1')

     bchildstep1_2 = st.checkbox('Sub step: child bundling step 1-2',value=True)
     if bchildstep1_2:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('Bundling Child Step 1-2')

  if bparentstep1warning:
        bps1warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: bps1)")
        p = document.add_paragraph(style=document.styles['List Bullet 1'])
        run = p.add_run(bps1warning)
        run.font.color.rgb = RGBColor(255, 0, 0)
        run.font.bold = True
     
  
  st.markdown("--------------------------")
  bparentstep2 = st.checkbox('Step: Parent Bundling Step 2',value=True)
  bparentstep2warning = st.checkbox('Any warning regarding bundling step 2?')
  if bparentstep2:
     p = document.add_paragraph(style=document.styles['List Bullet 1'])
     p.add_run('Bundling Parent Step 2')
     st.caption('Choose specific bundling steps')
     
     bchildstep2_1 = st.checkbox('Sub step: child bundling step 2-1',value=True)
     if bchildstep2_1:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('Bundling Child Step 2-1')

     bchildstep2_2 = st.checkbox('Sub step: child bundling step 2-2',value=True)
     if bchildstep2_2:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('Bundling Child Step 2-2')

  if bparentstep2warning:
        bps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: bps2)")
        p = document.add_paragraph(style=document.styles['List Bullet 1'])
        run = p.add_run(bps2warning)  
        run.font.color.rgb = RGBColor(255, 0, 0)
        run.font.bold = True
  
#-------------------------------------------------------------------------------
if cartoning:
  st.markdown("--------------------------")
  st.markdown('<p style="font-size: 20px;">Cartoning is selected</p>', unsafe_allow_html=True)
  st.markdown("--------------------------")
  p = document.add_paragraph(style=document.styles['List Bullet 0'])
  p.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 24 points
  # Main Process Name
  p.add_run('Cartoning')

  cparentstep1 = st.checkbox('Step: Parent cartoning Step 1',value=True)  
  cparentstep1warning = st.checkbox('Any warning regarding cartoning step 1?')
  if cparentstep1:
     p = document.add_paragraph(style=document.styles['List Bullet 1'])
     p.add_run('cartoning Parent Step 1')
     st.caption('- Choose specific bundling steps')
     
     cchildstep1_1 = st.checkbox('Sub step: child cartoning step 1-1',value=True)
     if cchildstep1_1:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('cartoning Child Step 1-1')

     cchildstep1_2 = st.checkbox('Sub step: child cartoning step 1-2',value=True)
     if cchildstep1_2:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('cartoning Child Step 1-2')
  
  if cparentstep1warning:
        cps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: cps1)")
        p = document.add_paragraph(style=document.styles['List Bullet 1'])
        run = p.add_run(cps2warning) 
        run.font.color.rgb = RGBColor(255, 0, 0)
        run.font.bold = True

  st.markdown("--------------------------")
  cparentstep2 = st.checkbox('Step: Parent cartoning Step 2',value=True)
  cparentstep2warning = st.checkbox('Any warning regarding cartoning step 2?')
  if cparentstep2:
     p = document.add_paragraph(style=document.styles['List Bullet 1'])
     p.add_run('cartoning Parent Step 2')
     st.caption('- Choose specific bundling steps')

     cchildstep2_1 = st.checkbox('Sub step: child cartoning step 2-1',value=True)
     if cchildstep2_1:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('cartoning Child Step 2-1')

     cchildstep2_2 = st.checkbox('Sub step: child cartoning step 2-2',value=True)
     if cchildstep2_2:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('cartoning Child Step 2-2')

  if cparentstep2warning:
        cps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: cps2)")
        p = document.add_paragraph(style=document.styles['List Bullet 1'])
        run = p.add_run(cps2warning) 
        run.font.color.rgb = RGBColor(255, 0, 0)
        run.font.bold = True

#-------------------------------------------------------------------------------

if additional:
  st.markdown("--------------------------")
  st.markdown('<p style="font-size: 20px;">Additional is selected</p>', unsafe_allow_html=True)
  st.markdown("--------------------------")
  p = document.add_paragraph(style=document.styles['List Bullet 0'])
  p.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 24 points
  # Main Process Name
  p.add_run('Additional')

  aparentstep1 = st.checkbox('Step: Parent additional Step 1',value=True)
  aparentstep1warning = st.checkbox('Any warning regarding additional step 1?')
  if aparentstep1:
     p = document.add_paragraph(style=document.styles['List Bullet 1'])
     p.add_run('additional Parent Step 1')
     st.caption('- Choose specific bundling steps')
     
     achildstep1_1 = st.checkbox('Sub step: child additional step 1-1',value=True)
     if achildstep1_1:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('additional Child Step 1-1')

     achildstep1_2 = st.checkbox('Sub step: child additional step 1-2',value=True)
     if achildstep1_2:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('additional Child Step 1-2')
  
  if aparentstep1warning:
        aps1warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: aps1)")
        p = document.add_paragraph(style=document.styles['List Bullet 1'])
        run = p.add_run(aps1warning) 
        run.font.color.rgb = RGBColor(255, 0, 0)
        run.font.bold = True

  st.markdown("--------------------------")
  aparentstep2 = st.checkbox('Step: Parent additional Step 2',value=True)
  aparentstep2warning = st.checkbox('Any warning regarding additional step 2?')
  if aparentstep2:
     p = document.add_paragraph(style=document.styles['List Bullet 1'])
     p.add_run('additional Parent Step 2')
     st.caption('- Choose specific bundling steps')

     achildstep2_1 = st.checkbox('Sub step: child additional step 2-1',value=True)
     if achildstep2_1:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('additional Child Step 2-1')

     achildstep2_2 = st.checkbox('Sub step: child additional step 2-2',value=True)
     if achildstep2_2:
        p = document.add_paragraph(style=document.styles['List Bullet 2'])
        p.add_run('additional Child Step 2-2')
  
  if aparentstep2warning:
        aps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: aps2)")
        p = document.add_paragraph(style=document.styles['List Bullet 1'])
        run = p.add_run(aps2warning)
        run.font.color.rgb = RGBColor(255, 0, 0)
        run.font.bold = True

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
