# **PBCREATOR**

###### Tool for generating Word documents based on Excel file content and also print documents in PDF format. We´re using to generate Playbooks documents <sup>with gui elements</sup>

#### Prerequisites

+ MS Windows
+ MS Office
+ Python
+ Python Requirements modules
* Windows OS

#### Installation

pip install -r requirements.txt

### How to use

1. Create docx template with jija2 style variables(`{{ variable name }}`)

2. Create Excel file where one sheet will contain all necessary information.
   - First line should contain variables names.
   - Second line can be description.
   - Columns where no variable is set are ignored by the script.

3. Configure conf file in the same directory as PBCreator executable
   - docx_templates_count        - number of templates to process
   - docx_template_name_1        - name of template 1 
   - docx_template_naming_1      - format for resulting document name for template 1 (see usage of variables in examples)
   - docx_template_dir_1         - initial directory for open template dialog for template 1
   - source_file_dir             - initial directory for open Excel file dialog
   - complete_data_sheet_name    - name of sheet in Excel file which should be processed
   - lines_to_skip               - number of lines used for description(yellow color in examples)
   - var_to_indicate_row         - variable to indicate row as filled 

4. Choose the output directory

### Execute

**** Change the first column of the Excel Spreadsheet file to X to create your Playbook Document otherwise leave empty.

python pbcreator.py

   ██████╗░██████╗░░█████╗░██████╗░███████╗░█████╗░████████╗░█████╗░██████╗░
██╔══██╗██╔══██╗██╔══██╗██╔══██╗██╔════╝██╔══██╗╚══██╔══╝██╔══██╗██╔══██╗
██████╔╝██████╦╝██║░░╚═╝██████╔╝█████╗░░███████║░░░██║░░░██║░░██║██████╔╝
██╔═══╝░██╔══██╗██║░░██╗██╔══██╗██╔══╝░░██╔══██║░░░██║░░░██║░░██║██╔══██╗
██║░░░░░██████╦╝╚█████╔╝██║░░██║███████╗██║░░██║░░░██║░░░╚█████╔╝██║░░██║
╚═╝░░░░░╚═════╝░░╚════╝░╚═╝░░╚═╝╚══════╝╚═╝░░╚═╝░░░╚═╝░░░░╚════╝░╚═╝░░╚═╝

made by: Logicalis CDOC!                           version 1.0

1 -- LOGICALIS
2 -- BRASILSEG
3 -- FESP
4 -- ACH
5 -- ONNET
6 -- Exit
Enter your choice:

LOGICALIS option will create Playbooks docs for the client id "BR-01";
BRASILSEG option will create Playbooks docs for the client id "BR-05";
FESP option will create Playbooks docs for the client id "BR-02";
ACH option will create Playbooks docs for this client id "CO-03";
ONNET option will create Playbooks docs for this client id "CO-06";

## 1.2 Changes

->Add Info Playbook Template;
->Change code to run INFORMATIONAL Playbooks with different templates.

## Knows Issues

-> Remove the "" and <> character of the Excel Database in order the script run correctly.

-> Sometimes the Microsft Word process (word.exe) and Microsft Excel process (excel.exe) freezes, producing some errors with the script, you need to close these orphan process in the Process Manager of Windows in order to work.

-> Name conflict with _FilterDatabase, the solution is clean up your Windows TEMP files.
