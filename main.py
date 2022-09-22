import pygsheets
import pandas as pd

""" Create and edit Word Document's DocProperties based on user input & Google Sheet Data
"""


def main():

    # authorization
    gc = pygsheets.authorize(service_file='credentials.json')
    file =

    # open the google spreadsheet (where 'PY to Gsheet Test' is the name of my sheet)
    sh = gc.open('Sahla Smart Solutions')

    # select the first sheet
    wks = sh[0]

    cell_matrix = wks.get_all_values(returnas='matrix')
    # Column Structure: Project Code, Project Name, Project Description, Date Started, Date ended, Status,
    # Project Leader, Client, Contacts, Budget, Payment Status, Notes


if __name__ == '__main__':
    main()