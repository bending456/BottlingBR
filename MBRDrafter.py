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


##########################################
#----------------------------------------#
#----------------------------------------#
##########################################

st.header("Master Batch Record Drafter")
st.caption("Note: We may need to split primary and secondary options")

##########################################
#----------------------------------------#
#----------------------------------------#
##########################################

st.markdown("### ***README***")
with st.expander("User Guide",expanded=True):
   st.caption("Step 0: Check the box to prevent auto reset")
   st.caption("Step 1: Type Document Name in Sidebar ")
   st.caption("Step 2: Type Output Name in Sidebar ")
   st.caption("Step 3: Type Batch Number in Sidebar ")
   st.caption("Step 4: Type Your Name in Sidebar ")
   st.caption("Step 5: Select Packaging Process(s) - [List of Processes] ")
   st.caption("Step 6: Select Specific Steps - [Process Control Panel] ")
   st.caption("Step 7: Download docx file")

if 'writing draft' not in st.session_state:
   st.session_state['writing draft']=False

document = Document()

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




##########################################
#----------------------------------------#
#----------------------------------------#
##########################################


###------ Title of Document
title = document.add_paragraph()
titleText = st.sidebar.text_input("Step 1: Write the title of your document in here ... (ex: Bundling procedure for XXX project)")
outputfileName = st.sidebar.text_input("Step 2: Write the name of output (docx file) name in here ...")
run = title.add_run(titleText)
run.bold = True
run.font.size = Pt(16)

###------ Header of Document - Batch Number
header1 = document.sections[0].header
paragraph = header1.paragraphs[0]
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
batchnumber = st.sidebar.text_input("Step 3: Write the batch number in here ...")
run = paragraph.add_run(batchnumber)
run.bold = True
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(255, 0, 0)
paragraph.paragraph_format.space_before=Pt(0)
paragraph.paragraph_format.space_after=Pt(0)

###------ Header of Document - Name
authorname = st.sidebar.text_input("Step 4: Write the name of author in here ...")
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


##########################################
#----------------------------------------#
#----------------------------------------#
##########################################

##---- Selecting Processes
st.markdown("### Step 5: List of Processes")
col1, col2 = st.columns(2)

### -------------- Primary Packaging -------------------

with col1:
   st.markdown('#### Primary Packaging')
   primary = st.checkbox("Primary Packaging")

   # Define all checkbox variables first
   sachet = canister = cotton = additional1 = False
   if primary:
      ###--- Primary Packaging related list
      sachet = st.checkbox("Sachet?")
      canister = st.checkbox("Canister? - N/A")
      cotton = st.checkbox("Cotton Filler? - N/A")
      additional1 = st.checkbox("Additional1? - N/A")

      subtitle = document.add_paragraph()
      run = subtitle.add_run('Primary Packaging')
      run.bold = True
      run.font.size = Pt(14)
      i = 0


##------------ Control Panel -----------------------------
st.divider()
st.markdown('### Step 6: Process Control Panel')


##################################################################
if not primary:
   st.caption("Primary is not Selected")
   st.divider()

if primary:
   st.markdown('#### Primary Packaging Step Selection')

if sachet:
   with st.expander('Select Steps for Sachet Process',expanded=True):
      p = document.add_paragraph(style=document.styles['List Bullet 0'])
      p.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 24 points
      # Main Process Name
      run = p.add_run('Sachet Process')
      run.bold = True
      run.font.size = Pt(12)
      sparentstep1 = st.checkbox('Step: Parent Sachet Step 1',value=True)
      if sparentstep1:
         p = document.add_paragraph(style=document.styles['List Bullet 1'])
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
            p = document.add_paragraph(style=document.styles['List Bullet 1'])
            run = p.add_run(xps1warning)
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True
   
      st.divider()
      sparentstep2 = st.checkbox('Step: Parent Sachet Step 2',value=True)
      if sparentstep2:
         p = document.add_paragraph(style=document.styles['List Bullet 1'])
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
            p = document.add_paragraph(style=document.styles['List Bullet 1'])
            run = p.add_run(xps2warning)  
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True

#################################################################################


## ---- Secondary Packaging ------------------------------

with col2:
   st.markdown('#### Secondary Packaging')
   secondary = st.checkbox("Secondary Packaging")
   
   cartoning = sidesert = bundling = shipper = additional2 = False
   
   if secondary:
      ###--- Secondary Packaging related list
      cartoning = st.checkbox("Cartoning?")
      sidesert = st.checkbox("Sidesert?")
      bundling = st.checkbox("Bundling?")
      shipper = st.checkbox("Shipper?")
      additional2 = st.checkbox("Additional2?")

      subtitle = document.add_paragraph()
      run = subtitle.add_run('Secondary Packaging')
      run.bold = True
      run.font.size = Pt(14)


########################################################################
if not secondary:
   st.caption("Secondary is not Selected")
   st.divider()

if secondary:
   st.markdown('#### Secondary Packaging Step Selection')

#- ------------------------------------------------------------------------------
if sidesert:
   with st.expander('Select Steps for Sidesert Process',expanded=True):
      p = document.add_paragraph(style=document.styles['List Bullet 0'])
      p.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 24 points
      # Main Process Name
      run = p.add_run('Adding Sidesert')
      run.bold = True
      run.font.size = Pt(12)
      ssparentstep1 = st.checkbox('Step 1: Preparing Sidesert',value=True)  
      
      if ssparentstep1:
         p = document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run('Weighing Sidesert')
         st.caption('- Choose specific weighing steps')
         sschildstep1_1 = st.checkbox('Step 1-A: Collect 10 sideserts and printweigh in the space provided. Record the scale number and lot number in the spaces provided.',value=True)
         if sschildstep1_1:
            p = document.add_paragraph(style=document.styles['List Bullet 2'])
            p.add_run('Collect 10 sideserts and printweigh in the space provided. Record the scale number and lot number in the spaces provided. \nRecord the sidesert usage log on pages XX-XX')
            p= document.add_paragraph(style=document.styles['List Bullet 3'])
            p.add_run("New Column 1: Scale #")
            p= document.add_paragraph(style=document.styles['List Bullet 3'])
            p.add_run("New Column 2: Lot #")
            p= document.add_paragraph(style=document.styles['List Bullet 3'])
            p.add_run("New Column 3: blank to print the weight")
            p = document.add_paragraph(style=document.styles['List Bullet 2'])
            p.add_run('Use the following calculation to determine the average weight of one sidesert')
            p = document.add_paragraph(style=document.styles['List Bullet 3'])
            p.add_run('__________ g (Wt. of 10 sidesert) / 10 = __________ g (Avg. Wt. of one sidesert)')
      
      ssparentstep1warning = st.checkbox('Any warning regarding adding sidesert step 1?')
      if ssparentstep1warning:
            ssps1warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: ssps1)")
            p = document.add_paragraph(style=document.styles['List Bullet 1'])
            run = p.add_run(ssps1warning) 
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True

      st.divider()
      ssparentstep2 = st.checkbox('Step 2: Preparing Sidesert',value=True)
      if ssparentstep2:
         p = document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run('Once the labeler machine is et up, remove 5 sideserts from the sidesert stream. Using maker, draw a line diagonally through the center of each sidesert. Apply those sideserts to the bottle and place them back. Ensure each bottle is rejected. Circle pass or fail')
         run = p.add_run('If the sideserts are not rejected, stop and contact a Supervisor or above to perform any adjustments needed')
         run.font.bold = True
         p = document.add_paragraph(style=document.styles['List Bullet 2'])
         p.add_run('New Column: Circle one Pass or Fail')
         p = document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run('Place the sideserts on the machine channel to ensure that the sideserts are facing the correct way. Circle pass or fail')
         run = p.add_run('If the sideserts are not rejected, stop and contact a Supervisor or above to perform any adjustments needed. \nNote: This is to ensure that the barcode is facing out. Once placed on the bottle, the barcode is facing out and detectable.')
         run.font.bold = True
         p = document.add_paragraph(style=document.styles['List Bullet 2'])
         p.add_run('New Column: Circle one Pass or Fail')

      ssparentstep2warning = st.checkbox('Any warning regarding adding sidesert step 2?')
      if ssparentstep2warning:
            ssps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: ssps2)")
            p = document.add_paragraph(style=document.styles['List Bullet 1'])
            run = p.add_run(ssps2warning) 
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True

#- ------------------------------------------------------------------------------
if cartoning:
   with st.expander('Select Steps for Cartoning Process',expanded=True):
      p = document.add_paragraph(style=document.styles['List Bullet 0'])
      p.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 24 points
      # Main Process Name
      run = p.add_run('Cartoning')
      run.bold = True
      run.font.size = Pt(12)
      cparentstep1 = st.checkbox('Step 1: Preparing Cartons',value=True)  
      
      if cparentstep1:
         p = document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run('Weighing Cartons')
         st.caption('- Choose specific weighing steps')
         cchildstep1_1 = st.checkbox('Step 1-A: Collect 10 cartons and printweigh in the space provided. Record the scale number and lot number in the spaces provided.',value=True)
         if cchildstep1_1:
            p = document.add_paragraph(style=document.styles['List Bullet 2'])
            p.add_run('Collect 10 cartons and printweigh in the space provided. Record the scale number and lot number in the spaces provided. \nRecord the carton usage log on pages XX-XX')
            p= document.add_paragraph(style=document.styles['List Bullet 3'])
            p.add_run("New Column 1: Scale #")
            p= document.add_paragraph(style=document.styles['List Bullet 3'])
            p.add_run("New Column 2: Lot #")
            p= document.add_paragraph(style=document.styles['List Bullet 3'])
            p.add_run("New Column 3: blank to print the weight")
            p = document.add_paragraph(style=document.styles['List Bullet 2'])
            p.add_run('Use the following calculation to determine the average weight of one carton')
            p = document.add_paragraph(style=document.styles['List Bullet 3'])
            p.add_run('__________ g (Wt. of 10 cartons) / 10 = __________ g (Avg. Wt. of one carton)')
      cparentstep1warning = st.checkbox('Any warning regarding cartoning step 1?')
      if cparentstep1warning:
            cps1warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: cps1)")
            p = document.add_paragraph(style=document.styles['List Bullet 1'])
            run = p.add_run(cps1warning) 
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True

      st.divider()
      cparentstep2 = st.checkbox('Step 2: Preparing Cartoner',value=True)
      if cparentstep2:
         p = document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run('Set up all cartoner infeed and outfeed conveyors to match the bottle and carton ins use')
         p = document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run('Set up the cartoner in the Dry Run mode and allow to cycle for NLT 1 minute.\nVerify a smooth cycle')
         p = document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run('Allow NLT 5 bottles to be loaded, formed, filled and sealed by turning Dry Run OFF.')
         p = document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run('Once NLT 5 bottles have been loaded, turn Dry Run back On.')
         p = document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run('Remove NLT 5 completed cartons from the exit conveyor ahead of CartonTracker for inspection. Indicate in the space provided if inspection is a PAss or FAil. If any failures are found, contact a Supervisor or above to perform any adjustment as needed')
         p = document.add_paragraph(style=document.styles['List Bullet 2'])
         p.add_run('New Column: Inspection Results (Circle One) Pass or Fail')
         p= document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run("Verify that the following are correct (carton will need to be opened for some items):")
         p= document.add_paragraph(style=document.styles['List Bullet 2'])
         p.add_run("Carton is properly closed and sealed")
         p= document.add_paragraph(style=document.styles['List Bullet 2'])
         p.add_run("Bottle is present and properly oriented")
         p= document.add_paragraph(style=document.styles['List Bullet 2'])
         p.add_run("Insert (if applicable) is present and properly oriented")
         p= document.add_paragraph(style=document.styles['List Bullet 2'])
         p.add_run("Any other required components (pill pack, etc.) are present and properly oriented")
         p= document.add_paragraph(style=document.styles['List Bullet 2'])
         p.add_run("External seals or labels are applied in correct location")
         p= document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run("Gather any reusable components (bottles, leaflets, pill packes, etc.) and return to appropriate location for rework")
         p= document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run("Ensure the inspected cartons are rejected. ")
         run = p.add_run("Note: Cartons cannot be reworked.")
         run.font.bold = True
         p= document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run("If seals or labels are applied to the carton, perform a challenge of the vision system")
         p= document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run("Pass 5 cartons through the camera system with NLT 1 seal/label missing per carton. Verify all 5 cartons are rejected. Indicate in the space provided if inspection is a Pass or Fail. If any failures are found contact a Supervisor or above to perform an adjustement as needed")
         p= document.add_paragraph(style=document.styles['List Bullet 2'])
         p.add_run("New Column: Inspection Results (Circle One) Pass or Fail")
         p= document.add_paragraph(style=document.styles['List Bullet 1'])
         p.add_run("Using the change over list, start setting each station to the correct setting. Then reinstall correct change parts. Once installed, go to machine configuration and press Link tab.")
         run = p.add_run("Note: Once machine is setup for processing, minor adjustment may be needed")
         run.font.bold = True

      cparentstep2warning = st.checkbox('Any warning regarding cartoning step 2?')
      if cparentstep2warning:
            cps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: cps2)")
            p = document.add_paragraph(style=document.styles['List Bullet 1'])
            run = p.add_run(cps2warning) 
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True

if bundling:  
   with st.expander('Select Steps for Bundling Process',expanded=True):
      p = document.add_paragraph(style=document.styles['List Bullet 0'])
      p.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 24 points
      # Main Process Name
      run = p.add_run('Bundling')
      run.bold = True
      run.font.size = Pt(12)
      bparentstep1 = st.checkbox('Step: Parent Bundling Step 1',value=True)

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
      
      bparentstep1warning = st.checkbox('Any warning regarding bundling step 1?')
      if bparentstep1warning:
            bps1warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: bps1)")
            p = document.add_paragraph(style=document.styles['List Bullet 1'])
            run = p.add_run(bps1warning)
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True
   
      st.divider()
      bparentstep2 = st.checkbox('Step: Parent Bundling Step 2',value=True)
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
      bparentstep2warning = st.checkbox('Any warning regarding bundling step 2?')
      if bparentstep2warning:
            bps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: bps2)")
            p = document.add_paragraph(style=document.styles['List Bullet 1'])
            run = p.add_run(bps2warning)  
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True


#-------------------------------------------------------------------------------
if additional2:
   with st.expander('Select Steps for Additional Process',expanded=True):
      p = document.add_paragraph(style=document.styles['List Bullet 0'])
      p.paragraph_format.line_spacing = Pt(10)  # Set line spacing to 24 points
      # Main Process Name
      run = p.add_run('Additional')
      run.bold = True
      run.font.size = Pt(12)
      aparentstep1 = st.checkbox('Step: Parent additional Step 1',value=True)

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
      
      aparentstep1warning = st.checkbox('Any warning regarding additional step 1?')
      if aparentstep1warning:
            aps1warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: aps1)")
            p = document.add_paragraph(style=document.styles['List Bullet 1'])
            run = p.add_run(aps1warning) 
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True

      st.divider()
      aparentstep2 = st.checkbox('Step: Parent additional Step 2',value=True)
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
      
      aparentstep2warning = st.checkbox('Any warning regarding additional step 2?')
      if aparentstep2warning:
            aps2warning = st.text_input("Please, explain the step that ops need to take extra caution (warning ID: aps2)")
            p = document.add_paragraph(style=document.styles['List Bullet 1'])
            run = p.add_run(aps2warning)
            run.font.color.rgb = RGBColor(255, 0, 0)
            run.font.bold = True


##########################################
#----------------------------------------#
#----------------------------------------#
##########################################

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
           file_name=outputfileName+".docx",
           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
       )

   # Remove temporary file
   os.unlink(tmp.name)
