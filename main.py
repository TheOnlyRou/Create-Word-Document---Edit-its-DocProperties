import pygsheets
import pandas as pd
import aspose.words as aw
from colorclass import Color

""" Create and edit Word Document's DocProperties based on user input & Google Sheet Data

Project relies on the package pygsheet: https://github.com/nithinmurali/pygsheets
For an Alternate way that uses Oauth: https://developers.google.com/sheets/api/quickstart/python
Aspose.Words Word Processing API: https://github.com/aspose-words/Aspose.Words-for-Python-via-.NET

Prerequisites to use this project:
    - Google Cloud Platform Project
    - Activate Google Drive & Google Sheets API on the GCP project
    - GCP Service Account for the project
         - JSON Service Account key for the service account (Get from IAm & Admin -> Service Accounts)    
    - Share your Google Sheets spreadsheet with your Service account
    - 
"""

# No. of implemented actions
IMP_ACTIONS = 2

# Project Codes Data Containers
proj_code_prefixes = []
proj_code_id_last = {}
proj_code_client = {}

proj_codes_all = []
proj_codes_id_counters = {}

# Document Template Titles Data Containers
template_title_code = {}
template_title_type = {}

# List of used custom document properties
custom_props = []


def main():
    # Reading from args txt file
    file = "args.txt"
    f = open(file, "r")

    credentials_path = f.readline()
    credentials_path = credentials_path.replace("\n", "")
    templates_dir = f.readline()
    templates_dir = templates_dir.replace("\n", "")
    options = [line.replace("\n", "") for line in f.readlines() if line.strip()]

    # authorization with GCP
    gc = pygsheets.authorize(service_file=credentials_path)

    # Get allowed Sheets titles & IDs
    allow_sheets = gc.spreadsheet_ids()
    allow_sheets_titles = gc.spreadsheet_titles()
    print(f'Sheet Titles: {allow_sheets_titles}')

    # open the google spreadsheet by its title
    sh = gc.open(allow_sheets_titles[0])

    # select the worksheets you want to manipulate
    sheet1 = sh[0]
    sheet2 = sh[2]

    # Return the whole worksheet as a matrix
    sheet1_matrix = sheet1.get_all_values(returnas='matrix')
    sheet2_matrix = sheet2.get_all_values(returnas='matrix')

    # Parse the data from the worksheets
    parse_data_projects(sheet1_matrix)
    parse_data_templates(sheet2_matrix)

    actions_index = len(options)
    selected_action = prompt_user_entry(options)

    if selected_action > IMP_ACTIONS or selected_action < 1:
        print(Color("{red}This action is not implemented yet. Please select another action{/red}"))
        selected_action = prompt_user_entry(options)

    # elif selected_action > actions_index
    #     print(Color("{red}This action is not implemented yet. Please select another action{/red}"))
    #     selected_action = prompt_user_entry()

    match selected_action:
        case 1:
            create_document_from_template()
        case 2:
            edit_existing_document()
        case 3:
            create_new_template()
        case _:
            print(Color("{red}This action is not implemented yet. Please select another action{/red}"))
            selected_action = prompt_user_entry(options)


def prompt_user_entry(options):
    """
     Prompt user to enter an integer corresponding to one of the actions written in the args file
    :param options: list
    contains all possible options to be presented to the user
    :return:
    """
    slct = -1
    print(Color("{red}Please select from the list below\n{/red}"))
    for option in options:
        print(f"{options.index(option)+1}. {option}")
    slct_option = input(Color("\n{green}Your Selection{/green}")+":\n")
    try:
        slct = int(slct_option)
    except:
        print(f'Your entry {slct_option} is incorrect. Please enter a number corresponding to an action from the list!'
              f'\nTry again ...')
        slct = prompt_user_entry(options)
        return slct
    return slct


def create_document_from_template():
    """ Copy template and copy data from spreadsheet referencing project if project report/quotation/proposal

    Prompts user to edit
    :return:
    """
    # TODO complete the method functionality
    print("CREATE NEW DOCUMENT FROM TEMPLATE SELECTED. ACTION NOT IMPLEMENTED YET")
    options = list(template_title_code.keys())
    slct = prompt_user_entry(options)
    try:
        print(f"Template file: {list(template_title_code.values())[slct]}.docx")
        doc_path = "templates\\"+list(template_title_code.values())[slct]+".docx"
        doc = aw.Document(doc_path)
        doc_name = input(Color("{green}Name of the document:{/green}:\n"))
        while doc_name[0].isdigit():
            print(Color("{red}File names can't start with a digit! try agan{/red}"))
            doc_name = input(Color("{green}Name of the document:{/green}:\n"))
        doc_name = doc_name + "(" + list(template_title_code.keys())[slct] + ")"
        print(Color("{green}Please enter the following document details:{/green}:\n"))
        my_props = doc.custom_document_properties
        for prop in my_props:
            if prop in custom_props:
                edited_field_val = input(f"{prop.name} ({prop.type}):\n")
                prop.value = edited_field_val
        try:
            # TODO prompt user for where to save file
            doc.save(doc_name+".docx")
        except IOError:
            print("Unable to save file. Retry or contact the IT team if the issue persists")
    except:
        print("Error Occurred. Retry or contact the IT team if the issue persists")


def edit_existing_document():
    """ Locate a file and edit its properties

    :return:
    """
    doc_path = input(Color("{green}Document Path{/green}:\n"))
    try:
        doc = aw.Document(doc_path)
        my_props = doc.custom_document_properties
        for prop in my_props:
            if prop in custom_props:
                print(f"{custom_props.index(prop.name)+1} : {prop.name} ({prop.type}):\n")
        match input(Color("{green}Do you wish to edit any of these properties?(Y/N){/green}:\n")):
            case ["y" | "Y"]:
                try:
                    slct = int(input(Color("{green}Enter index of the property you want to edit{/green}:\n")))
                    if not custom_props[slct]:
                        print("You need to enter an integer from the list of properties! Try again")
                        edit_existing_document()
                    my_props.add(custom_props[slct], input(custom_props[slct]))
                except:
                    print("You need to enter an integer from the list of properties! Try again")
                    edit_existing_document()
    except:
        print("File not found. Please try again!")
        edit_existing_document()


def create_new_template():
    # TODO
    """Copy Template of templates, create new document using it, using type and give it a new id, then add it to the
    templates worksheet

    :return:
    """

    pass


def parse_data_templates(sheet2_matrix):
    """ Parse Template data and populate global doc template data containers

    :param sheet2_matrix: list[list]
    :return:
    """
    for line in sheet2_matrix:
        if sheet2_matrix.index(line) == 0:
            print(f"LOG: Line {sheet2_matrix.index(line)} Skipped")
            continue
        if line[0] == "":
            print(f"LOG: Nothing to track after line {sheet2_matrix.index(line)}. Stopping ...")
            break
        else:
            template_title_code[line[0]] = line[3]
            template_title_type[line[0]] = line[1]
    print(f"Template Titles & Codes: {template_title_code}")
    print(f"Template Titles & Types: {template_title_type}")


def parse_data_projects(sheet1_matrix):
    """
    Parse Project data and populate global project data containers
    :param sheet1_matrix: list[list]
    :return:
    """
    for line in sheet1_matrix:
        if sheet1_matrix.index(line) == 0:
            print(f"LOG: Line {sheet1_matrix.index(line)} Skipped")
            continue
        if sheet1_matrix.index(line) > 9 and line[0] == "":
            print(f"LOG: Nothing to track after line {sheet1_matrix.index(line)}. Stopping ...")
            break
        else:
            # Project Codes Prefixes
            if line[0] != "":
                print(f"LOG: PRJ_CODE: {line[0]} appended to Project Codes")
                proj_codes_all.append(line[0])
            if line[15] != "":
                print(f"LOG: PRJ_CODE_PRE {line[15]} appended to Project Prefixes")
                proj_code_prefixes.append(line[15])
    for prefix in proj_code_prefixes:
        proj_codes_id_counters[prefix] = 0

    for code in proj_codes_all:
        for prefix in proj_codes_id_counters.keys():
            if code.startswith(prefix):
                proj_codes_id_counters[prefix] += 1
                break
    print(f"Prefixes Count: {proj_codes_id_counters}")


if __name__ == '__main__':
    main()
