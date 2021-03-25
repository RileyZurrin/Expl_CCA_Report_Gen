# Expl_CCA_Report_Gen
A script to generate a report for exploratory CCA output.

Steps to run:

i) Add canoload, CCA, and PCA files from exploratory CCA output to your py project folder
    e.g.'canoload_RSPM_con.xlsx', 'CCA_RSPM_con.csv', and 'PCA_RSPM_con.csv'

ii) Make file labelled Variables.txt containing all behavioural variables seperated by commas (no spaces!)
    e.g. Age_at_qualt,EducationEstYears,VarGC
    
iii) Add Variables.txt file to py project folder

iv) Open ReportGen.py

v) Make sure the following packages are installed:

    a) numpy
    b) python-docx
    c) openpyxl
    d) pandas
  
vi) Modify the following lines:

      a) Line 11 with task abbreviation. e.g. Task = "RSPM"
      b) Lines 14 through 16 with aforementioned files names
      c) Line 19 with full network names (from classification). e.g. C1_Neg_92_TDMN_0.69 / PP_NegL92_AAR_0.57
      d) Line 22 with shortened network names. e.g. DMN_AAR
 
vii) Run it! The docx should then be in your py project folder.

viii) Open docx file and add Table of Contents: References > Table of Contents > Contemporary

ix) Adjust "Summary of Study, Networks, Conditions, and Behavioural Data" Section with task information.

x) You're done :)

