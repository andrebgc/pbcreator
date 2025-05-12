# **PBCREATOR**

###### Tool for generating Word documents based on Excel file content, and also print documents in PDF format. We´re using to generate Playbooks documents with GUI elements

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

2. Create an Excel file where one sheet will contain all the necessary information.
   - The first line should contain variable names.
   - The second line can be a description.
   - Columns where no variable is set are ignored by the script.

3. Configure the conf file in the same directory as the PBCreator executable
   - docx_templates_count        - number of templates to process
   - docx_template_name_1        - name of template 1 
   - docx_template_naming_1      - format for resulting document name for template 1 (see usage of variables in examples)
   - docx_template_dir_1         - initial directory for open template dialog for template 1
   - source_file_dir             - initial directory for the open Excel file dialog
   - complete_data_sheet_name    - name of sheet in Excel file which should be processed
   - lines_to_skip               - number of lines used for description(yellow color in examples)
   - var_to_indicate_row         - variable to indicate row as filled 

4. Choose the output directory

### Execute

**** Change the first column of the Excel Spreadsheet file to X to create your Playbook Document; otherwise, leave it empty.

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

Client1 option will create Playbooks docs for the client ID "CL-01";
Client2 option will create Playbooks docs for the client ID "Cl-02";

## 1.2 Changes

->Add Info Playbook Template;
->Change code to run INFORMATIONAL Playbooks with different templates.

## Knows Issues

-> Remove the "" and <> characters of the Excel Database so the script runs correctly.

-> Sometimes the Microsoft Word process (word.exe) and Microsoft Excel process (excel.exe) freeze, producing some errors with the script; you need to close these orphan processes in the Process Manager of Windows in order to work.

-> Name conflict with _FilterDatabase, the solution is to clean up your Windows TEMP files.
