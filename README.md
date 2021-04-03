# Expl_CCA_Report_Gen
A script to generate a report for exploratory CCA output.

Steps to run:

i) Add canoload, CCA, and PCA files for each group from exploratory CCA output to your py project folder
    e.g.'canoload_LDT_Con.xlsx', 'CCA_LDT_Con.csv', 'PCA_LDT_Con.csv', 'canoload_LDT_Bip.xlsx', 'CCA_LDT_Bip.csv', 'PCA_LDT_Bip.csv'

ii) Make Variables.csv file containing all behavioural variables. The easiest way to do this is to open the input files (e.g. demographics_Controls.csv) for exploratory CCA and copy and paste the variables into an Excel sheet. 
    e.g. Handedness, Sex, Gender_Identity
    
iii) Add Variables.csv file to py project folder

iv) Open ReportGen.py

v) Make sure the following packages are installed:

    a) numpy
    b) python-docx
    c) openpyxl
    d) pandas
  
vi) Modify the following lines:

      a) Line 11 with task abbreviation. e.g. Task = "LDT"
      b) Line 13 with Group names. e.g. Groups = ["Controls", "Bipolar Patients"]
      b) Lines 16 through 18 with aforementioned files names
      c) Line 21 with full network names (from classification). e.g. C1_Neg_92_TDMN_0.69 / PP_NegL92_AAR_0.57
      d) Line 23 with shortened network names in same order as full network names. Make sure names are consistent with network names in canoload file. e.g. DMN_AAR
 
vii) Run it! The docx should then be in your py project folder.

viii) Open the docx file and add Table of Contents: References > Table of Contents

ix) Adjust "Summary of Study, Networks, Conditions, and Behavioural Data" Section with task information.

x) You're done :)

