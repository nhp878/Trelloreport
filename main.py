# -*- coding: utf-8 -*-
# Simple script to generate Excel report from extracted card info out of Trello
#
# TODO: Alldone
#
# How to run: from cmd type:
# python main.py "Sprint 2"
# or
# python main.py 2

import openpyxl
import requests
import sys
import time

# Creating an API Key and Token
# For the key, access this URL:
#   https://trello.com/1/appKey/generate
# For the token, access this:
#   https://trello.com/1/authorize?response_type=token&expiration=30days&name=gui&key=REPLACE_WITH_YOUR_KEY
API_KEY = '515adc1d50167dd117b1cdcb6c5fee72'
API_TOKEN = '3d22d2fde8bfb1f2b4e326a17609ba4d52ee4f14cf28515ea4c43584a64393db'

# Your Board ID
BOARD_ID = 'gobanyan'
DONE_LIST_ID = "62223e751556ad5e83fd87b4"
READY_REVIEW_LIST_ID = "62223e67fcc5f0191e9f9d94"


if __name__ == '__main__':
    # Create an empty dictionary
    total_dict = dict()
    label_list = list()
    sprint = str(sys.argv[1])
    if sprint.isnumeric():
        sprint = "sprint " + sprint

    #Create Excel file
    stt = 1
    column = 'A'
    theFile = openpyxl.load_workbook('template.xlsx')
    currentSheet = theFile['Sheet1']
    currentSheet['A1'].value = sprint.upper()
    try:
        cards = requests.get('https://trello.com/1/lists/%s/cards' % DONE_LIST_ID, params={'key': API_KEY, 'token': API_TOKEN}).json()
    except Exception as e:
        print("Oops!", e.__class__, "occurred.")
        print("Please run https://trello.com/1/authorize?response_type=token&expiration=30days&name=gui&key="
              "515adc1d50167dd117b1cdcb6c5fee72 to get new key and update in the application")
        exit(1)

    for card in cards:
        print(card['name'])
        cell_name = "{}{}".format('A', stt + 3)
        currentSheet[cell_name].value = str(stt)
        cell_name = "{}{}".format('B', stt + 3)
        currentSheet[cell_name].value = card['name']

        for label in card['labels']:
            if label['name'].lower() == sprint.lower():
                card_info = requests.get('https://trello.com/1/cards/%s/?fields=name&customFieldItems=true' % card['id'],
                                     params={'key': API_KEY, 'token': API_TOKEN}).json()
                customFields = card_info['customFieldItems']
                for customField in customFields:
                    field_info = requests.get('https://trello.com/1/customFields/%s?' % customField['idCustomField'], params={'key': API_KEY, 'token': API_TOKEN}).json()
                    field_name = field_info['name']
                    field_value = customField['value']['number']
                    print(field_name)
                    print(field_value)
                    current_column = ''
                    if field_name in total_dict:
                        total_dict[field_name] = total_dict[field_name] + int(field_value)
                        current_column = chr(label_list.index(field_name) + 67)

                    else:
                        total_dict[field_name] = int(field_value)
                        label_list.append(field_name)
                        current_column = chr(len(label_list) + 66)
                        cell_name = "{}{}".format(current_column, 3)
                        currentSheet[cell_name].value = field_name
                    cell_name = "{}{}".format(current_column, stt + 3)
                    currentSheet[cell_name].value = int(field_value)

        stt = stt + 1
    print(total_dict)

    #Print other titles
    cell_name = "{}{}".format('A', stt + 3)
    currentSheet[cell_name].value = 'TOTAL'

    for label in label_list:
        current_column = chr(label_list.index(label) + 67)
        cell_name = "{}{}".format(current_column, stt + 3)
        currentSheet[cell_name].value = total_dict[label]

    theFile.save(time.strftime("%Y%m%d-%H%M%S") + '_' + BOARD_ID + '_' + sprint + '.xlsx')
