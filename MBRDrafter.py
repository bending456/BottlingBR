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
from datetime import date


st.header("Master Batch Record Drafter [Under Construction]")
st.caption("Note: We may need to split primary and secondary options")
if 'writing draft' not in st.session_state:
   st.session_state['writing draft']=False

document = Document()

#---------- Sidebar Setup
stateholder = st.sidebar.checkbox("Check this box to prevent unwanted rerun")
if stateholder:
   st.session_state['writing draft']=True

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

###------ Title of Document
title = document.add_paragraph()
titleText = st.sidebar.text_input("Write the title of your document in here ... (ex: Bundling procedure for XXX project)")
outputfileName = st.sidebar.text_input("Write the name of output (docx file) name in here ...")
run = title.add_run(titleText)
run.bold = True
run.font.size = Pt(16)

###------ Header of Document - Batch Number
header1 = document.sections[0].header
paragraph = header1.paragraphs[0]
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
batchnumber = st.sidebar.text_input("Write the batch number in here ...")
run = paragraph.add_run(batchnumber)
run.bold = True
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(255, 0, 0)
paragraph.paragraph_format.space_before=Pt(0)
paragraph.paragraph_format.space_after=Pt(0)

###------ Header of Document - Name
authorname = st.sidebar.text_input("Write the name of author in here ...")
new_paragraph = header1.add_paragraph()
run2 = new_paragraph.add_run(authorname)
run2.bold = True
run2.font.size = Pt(10)
run2.font.color.rgb = RGBColor(0,0,255)
new_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
new_paragraph.paragraph_format.space_before=Pt(0)
new_paragraph.paragraph_format.space_after=Pt(0)

###------ Header of Document - Date
new_paragraph2 = header1.add_paragraph()
today = date.today()
formatted_date = today.strftime("%B %d, %Y")
run3 = new_paragraph2.add_run(f'{formatted_date}')
run3.bold = True
run3.font.size = Pt(10)
run3.font.color.rgb = RGBColor(0,0,0)
new_paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
new_paragraph2.paragraph_format.space_before=Pt(0)
new_paragraph2.paragraph_format.space_after=Pt(0)

###------ Footer of Document
footer = document.sections[0].footer
paragraph = footer.paragraphs[0]
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
run = paragraph.add_run("Confidential")
run.bold = True
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(0, 0, 255)

##---- Selecting Processes
st.markdown("### List of Processes")
col1, col2 = st.columns(2)

with col1:
   st.markdown('#### Primary Packaging')
   primary = st.checkbox("Primary Packaging")

   if primary:
      ###--- Primary Packaging related list
      sachet = st.checkbox("Sachet?")
      canister = st.checkbox("Canister?")
      cotton = st.checkbox("Cotton Filler?")
      additional1 = st.checkbox("Additional1?")

      subtitle = document.add_paragraph()
      run = subtitle.add_run('Primary Packaging')
      run.bold = True
      run.font.size = Pt(14)

with col2:
   st.markdown('#### Secondary Packaging')
   secondary = st.checkbox("Secondary Packaging")

   if secondary:
      ###--- Secondary Packaging related list
      bundling = st.checkbox("Bundling?")
      cartoning = st.checkbox("Cartoning?")
      additional2 = st.checkbox("Additional2?")

      subtitle = document.add_paragraph()
      run = subtitle.add_run('Secondary Packaging')
      run.bold = True
      run.font.size = Pt(14)


##------------ Control Panel -----------------------------
st.divider()
st.markdown('### Process Control Panel')
##################################################################
if not primary:
   st.caption("Primary is not Selected")
   st.divider()

elif primary and sachet:
   st.markdown('#### Primary Packaging Step Selection')
   with st.expander('Select Steps for Sachet Process',expanded=True):
      p = document.add_paragraph(style='List Number')
      p.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 24 points
      # Main Process Name
      run = p.add_run('Sachet Process')
      run.bold = True
      run.font.size = Pt(12)
      sparentstep1 = st.checkbox('Step: Parent Sachet Step 1',value=True)
      if sparentstep1:
         #p = document.add_paragraph(style='List Number 2')
         p = document.add_paragraph(style='List Number 2')
         p.add_run('Sachet Parent Step 1')
         st.caption('- Choose specific bundling steps')
         schildstep1_1 = st.checkbox('Sub step: child Sachet step 1-1',value=True)
         if schildstep1_1:
            p = document.add_paragraph(style=document.styles['List Bullet 2'])
            p.add_run('Sachet Child Step 1-1')
         
         schildstep1_2 = st.checkbox('Sub step: child Sachet step 1-2',value=True)
         if schildstep1_2:
            p = document.add_paragraph(style=document.styles['List Bullet 2'])
            p.add_run('Sachet Child Step 1-2')
      
      sparentstep1warning = st.checkbox('Any warning regarding Sachet step 1?')
      if sparentstep1warning:
            xps1warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: sps1)")
            #p = document.add_paragraph(style='List Number 2')
            p = document.add_paragraph(style='List Number 2')
            run = p.add_run(xps1warning)
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True
   
      st.divider()
      sparentstep2 = st.checkbox('Step: Parent Sachet Step 2',value=True)
      if sparentstep2:
         p = document.add_paragraph(style='List Number 2')
         p.add_run('Sachet Parent Step 2')
         st.caption('Choose specific Sachet steps')
         schildstep2_1 = st.checkbox('Sub step: child Sachet step 2-1',value=True)
         if schildstep2_1:
            p = document.add_paragraph(style=document.styles['List Bullet 2'])
            p.add_run('Sachet Child Step 2-1')
         schildstep2_2 = st.checkbox('Sub step: child Sachet step 2-2',value=True)
         if schildstep2_2:
            p = document.add_paragraph(style=document.styles['List Bullet 2'])
            p.add_run('Sachet Child Step 2-2')
      sparentstep2warning = st.checkbox('Any warning regarding Sachet step 2?')
      if sparentstep2warning:
            xps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: sps2)")
            p = document.add_paragraph(style='List Number 2')
            run = p.add_run(xps2warning)  
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True

#########################################################################
if not secondary:
   st.caption("Secondary is not Selected")
   st.divider()

elif bundling and secondary:
   st.markdown('#### Secondary Packaging Step Selection')
   with st.expander('Select Steps for Bundling Process',expanded=True):
      p = document.add_paragraph(style='List Number')
      p.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 24 points
      # Main Process Name
      run = p.add_run('Bundling')
      run.bold = True
      run.font.size = Pt(12)
      bparentstep1 = st.checkbox('Step: Parent Bundling Step 1',value=True)
      if bparentstep1:
         #p = document.add_paragraph(style='List Number 2')
         p = document.add_paragraph(style='List Number 2')
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
      
      bparentstep1warning = st.checkbox('Any warning regarding bundling step 1?')
      if bparentstep1warning:
            bps1warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: bps1)")
            #p = document.add_paragraph(style='List Number 2')
            p = document.add_paragraph(style='List Number 2')
            run = p.add_run(bps1warning)
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True
   
      st.divider()
      bparentstep2 = st.checkbox('Step: Parent Bundling Step 2',value=True)
      if bparentstep2:
         p = document.add_paragraph(style='List Number 2')
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
      bparentstep2warning = st.checkbox('Any warning regarding bundling step 2?')
      if bparentstep2warning:
            bps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: bps2)")
            p = document.add_paragraph(style='List Number 2')
            run = p.add_run(bps2warning)  
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True
   
#- ------------------------------------------------------------------------------
elif cartoning and secondary:
   st.markdown('#### Secondary Packaging Step Selection')
   with st.expander('Select Steps for Cartoning Process',expanded=True):
      p = document.add_paragraph(style='List Number')
      p.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 24 points
      # Main Process Name
      run = p.add_run('Cartoning')
      run.bold = True
      run.font.size = Pt(12)
      cparentstep1 = st.checkbox('Step: Parent cartoning Step 1',value=True)  
      
      if cparentstep1:
         p = document.add_paragraph(style='List Number 2')
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
      cparentstep1warning = st.checkbox('Any warning regarding cartoning step 1?')
      if cparentstep1warning:
            cps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: cps1)")
            p = document.add_paragraph(style='List Number 2')
            run = p.add_run(cps2warning) 
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True
      st.divider()
      cparentstep2 = st.checkbox('Step: Parent cartoning Step 2',value=True)
      if cparentstep2:
         p = document.add_paragraph(style='List Number 2')
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
      cparentstep2warning = st.checkbox('Any warning regarding cartoning step 2?')
      if cparentstep2warning:
            cps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: cps2)")
            p = document.add_paragraph(style='List Number 2')
            run = p.add_run(cps2warning) 
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True
#-------------------------------------------------------------------------------
elif additional2 and secondary:
   st.markdown('#### Secondary Packaging Step Selection')
   with st.expander('Select Steps for Additional Process',expanded=True):
      p = document.add_paragraph(style='List Number')
      p.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 24 points
      # Main Process Name
      run = p.add_run('Additional')
      run.bold = True
      run.font.size = Pt(12)
      aparentstep1 = st.checkbox('Step: Parent additional Step 1',value=True)
      if aparentstep1:
         p = document.add_paragraph(style='List Number 2')
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
      
      aparentstep1warning = st.checkbox('Any warning regarding additional step 1?')
      if aparentstep1warning:
            aps1warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: aps1)")
            p = document.add_paragraph(style='List Number 2')
            run = p.add_run(aps1warning) 
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True
      st.divider()
      aparentstep2 = st.checkbox('Step: Parent additional Step 2',value=True)
      if aparentstep2:
         p = document.add_paragraph(style='List Number 2')
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
      
      aparentstep2warning = st.checkbox('Any warning regarding additional step 2?')
      if aparentstep2warning:
            aps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: aps2)")
            p = document.add_paragraph(style='List Number 2')
            run = p.add_run(aps2warning)
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True

# Save the document
#document.save(outputfileName+'.docx')
st.sidebar.header("**Download Ready**")
if st.sidebar.checkbox("Check this box if the draft is ready"):
   with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
       document.save(tmp.name)
       tmp.seek(0)

       # Create a button to download the docx file
       st.sidebar.download_button(
           label="Download .docx file",
           data=tmp.read(),
           file_name=outputfileName+".docx",
           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
       )

   # Remove temporary file
   os.unlink(tmp.name)
