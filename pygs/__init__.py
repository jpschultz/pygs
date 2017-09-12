#!/usr/bin/env python

__author__ = "JP Schultz jp.schultz@gmail.com"
__license__ = "MIT"


import pandas as pd
import httplib2
import os
import string
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import pygs_tools as pytools
import initialize_service as initService


def create_empty_spreadsheet(document_name = None, sheet_name = None, **kwargs):
        
    """
    This will create an empty Google Sheet and return an object containing both the 'key' and the 'url'
    of the created sheet.
    
    Parameters
    ----------
    document_name : str, optional
        This will be the name of the Google Sheet. If left blank, it will default to 'Untitled Spreadsheet'
    
    sheet_name : str, optional
        This is the name of the first tab within the Google Sheet Doc. If not set, will be Sheet1.
        
    
    Returns
    -------
    Returns an object containing both the 'key' for the spreadsheet and the 'url'
    """
    service = initService._getService()
    
    cols = kwargs.pop('cols', 26)
    rows = kwargs.pop('rows', 1000)
    
    if document_name == None:
        document_name = 'Untitled spreadsheet'

    if sheet_name == None:
        sheet_name = 'Sheet1'

    sheets_info = [
            {
                'properties': {
                 'gridProperties':{
                    'columnCount': cols,
                    'rowCount': rows
                    },
                    'index': 0,
                    'sheetId': 0,
                    'sheetType': 'GRID',
                    'title': 'Sheet1'
                }
            }
        ]
    
    if document_name == None:
        response = service.spreadsheets().create(body={'sheets':sheet_info}).execute()
    else:
        response = service.spreadsheets().create(body={'properties':{'title':document_name}, 'sheets':sheets_info}).execute()   
    ret_val = {
                'spreadsheetId': str(response['spreadsheetId']),
                'spreadsheetUrl': str(response['spreadsheetUrl'])  
              }
        
    return ret_val


def create_spreadsheet_from_df(df, sheet_name = None, document_name = None, header=True):
    """
    Given a Pandas DataFrame (df), this will create a google sheet and name it the document_name and paste
    the DataFrame to the first sheet, which it will name the sheet_name
    
    Parameters
    ----------
    df : Pandas Dataframe
         This will take the dataframe and output it to a brand new Google Sheet  
    sheet_name : str, optional
        The name of the tab within the sheet that you would like the DataFrame pasted to. If left blank,
        will default to Sheet1
    document_name : str, optional
        This will be the name of the Google Sheet. If left blank, it will default to 'Untitled Spreadsheet'
    header : bool, optional
        This will determine if the header (column titles) are included when pasting data. Defaults to true.
        Setting to False will mean that just the raw data is pasted, starting in the first cell
    
    Returns
    -------
    Returns an object containing both the 'key' for the spreadsheet and the 'url'
        
    """
    if df.empty:
        raise ValueError('Please pass in a dataframe with data.')
    
    if header:
        paste_data = [df.columns.tolist()] + df.as_matrix().tolist()
    else:
        paste_data = df.as_matrix().tolist()
        
    total_cells = len(paste_data) * len(paste_data[0])
    
    if total_cells > 2000000:
        raise ValueError('There are more than 2 million cells in this dataframe which cannot be loaded into Google Sheets.')
    
    #clean/prep the dataframe to go to google sheets
    df = pytools._cleanDF(df)
    
    if document_name == None:
        document_name = 'Untitled spreadsheet'
    if sheet_name == None:
        sheet_name = 'Sheet1'
    
    last_col = pytools._getEndCol(paste_data)
    last_row = str(len(paste_data))
    
    a1notation = sheet_name + '!A1:' + last_col + last_row
    
    if len(paste_data) * 26 > 2000000:
        cols = len(paste_data[0])
        rows = len(paste_data)
    else:
        cols = 26
        rows = 1000
      
    
    service = initService._getService()
    
    new_sheet = create_empty_spreadsheet(document_name = document_name, sheet_name=sheet_name, cols=cols, rows=rows)
    new_sheet_id = new_sheet['spreadsheetId']
    
    body = {
        "range": a1notation,
        "values": paste_data
    }
    
    response = service.spreadsheets().values().update(spreadsheetId= new_sheet_id, range=a1notation, body=body, valueInputOption='USER_ENTERED').execute()
    

    ret_val = {
        'status' : 'success',
        'spreadsheetId': str(response['spreadsheetId']),
        'spreadsheetUrl': str(new_sheet['spreadsheetUrl'])  
    }

    return ret_val


def update_sheet_with_df(df, sheet_name, spreadsheetId, header=True):
    """
    Given a Pandas DataFrame (df), spreadsheetId and sheet_name, this will empty the sheet and paste the dataframe into it.
    
    Parameters
    ----------
    df : Pandas Dataframe, required
         This will take the dataframe paste it into a defined tab on a Google Sheet
    sheet_name : str, required
        The name of the tab within the sheet that you would like the DataFrame pasted to.
    spreadsheetId : str, required
        The ID of the spreadsheet containing the sheet that you want the DataFrame pasted to.
    header : bool, optional
        This will determine if the header (column titles) are included when pasting data. Defaults to true.
        Setting to False will mean that just the raw data is pasted, starting in the first cell
    
    Returns
    -------
    Returns an object containing both the 'key' for the spreadsheet and the 'url'
        
    """
    if df.empty:
        raise ValueError('Please pass in a dataframe with data.')
    if not sheet_name:
        raise ValueError('Please specify a sheet name.')
    if not spreadsheetId:
        raise ValueError('Please specify a spreadsheetId.')
    
    #clean/prep the dataframe to go to google sheets
    df = pytools._cleanDF(df)

    if header:
        paste_data = [df.columns.tolist()] + df.as_matrix().tolist()
    else:
        paste_data = df.as_matrix().tolist()
        
    total_cells = len(paste_data) * len(paste_data[0])
    
    if total_cells > 2000000:
        raise ValueError('There are more than 2 million cells in this dataframe which cannot be loaded into Google Sheets.')
    
    last_col = pytools._getEndCol(paste_data)
    last_row = str(len(paste_data))
    
    a1notation = sheet_name + '!A1:' + last_col + last_row
    
    service = initService._getService()
    
    current_state = service.spreadsheets().get(spreadsheetId = spreadsheetId).execute()
    
    current_cols = 26
    current_rows = 1000
    
    found = False
    for sheet in current_state['sheets']:
        if sheet['properties']['title'] == sheet_name:
            current_cols = sheet['properties']['gridProperties']['columnCount']
            current_rows = sheet['properties']['gridProperties']['rowCount']
            sheet_id = sheet['properties']['sheetId']
            found = True
            break

    if not found:
        raise ValueError("Unable to find '{}' in the spreadsheet. Please check the sheet name again.".format(sheet_name))

    #clear the data from the sheet
    clear_range_end_col = pytools._getEndCol([range(current_cols)])
    service.spreadsheets().values().clear(spreadsheetId=spreadsheetId, range=sheet_name+"!A1:" + clear_range_end_col + str(current_rows), body={}).execute()

    #remove extra columns so we can fit the new DF
    if (len(paste_data) * current_cols) > 2000000:
        body = {
        "requests":[{        
            "deleteDimension":
                {
            "range":
                    {
                    'sheetId':sheet_id,
                    'dimension':'COLUMNS',
                    'startIndex':1,
                    'endIndex':current_cols - 1
                    }
                }
            }]
        }
        #clear the columns
        service.spreadsheets().batchUpdate(spreadsheetId = spreadsheetId, body=body).execute()
    
    if len(paste_data) < current_rows:
        body = {
        "requests":[{        
            "deleteDimension":
                {
            "range":
                    {
                    'sheetId':sheet_id,
                    'dimension':'ROWS',
                    'startIndex':len(paste_data)-1,
                    'endIndex':current_rows
                    }
                }
            }]
        }
        #clear the columns
        service.spreadsheets().batchUpdate(spreadsheetId = spreadsheetId, body=body).execute()

    body = {
        "range": a1notation,
        "values": paste_data
    }
    response = service.spreadsheets().values().update(spreadsheetId=spreadsheetId, range=a1notation, body=body, valueInputOption='USER_ENTERED').execute()
    

    ret_val = {
        'status' : 'success',
        'spreadsheetId': str(response['spreadsheetId']),
        'spreadsheetUrl': 'https://docs.google.com/spreadsheets/d/' + str(response['spreadsheetId'])  
    }
        
    return ret_val


def create_tab_from_df(df, sheet_name, spreadsheetId, header=True):
    """
    Given a Pandas DataFrame (df), spreadsheetId and sheet name, this will empty the sheet and paste the dataframe into it.
    
    Parameters
    ----------
    df : Pandas Dataframe, required
         This will take the dataframe paste it into a defined tab on a Google Sheet
    sheet_name : str, required
        The name of the tab within the sheet that you want created.
    spreadsheetId : str, required
        The ID of the spreadsheet containing the sheet that you want the DataFrame pasted to.
    header : bool, optional
        This will determine if the header (column titles) are included when pasting data. Defaults to true.
        Setting to False will mean that just the raw data is pasted, starting in the first cell
    
    Returns
    -------
    Returns an object containing both the 'key' for the spreadsheet and the 'url'
        
    """
    if df.empty:
        raise ValueError('Please pass in a dataframe with data.')
    if not sheet_name:
        raise ValueError('Please specify a sheet name.')
    if not spreadsheetId:
        raise ValueError('Please specify a spreadsheetId.')
    
    #clean/prep the dataframe to go to google sheets
    df = pytools._cleanDF(df)

    if header:
        paste_data = [df.columns.tolist()] + df.as_matrix().tolist()
    else:
        paste_data = df.as_matrix().tolist()
    
    sheet_name = pytools._cleanSheetName(sheet_name, spreadsheetId)
    
    body = {
    "requests":[{        
        "addSheet":{
        "properties":{
            "title": sheet_name,
            "gridProperties":{
                "rowCount": len(paste_data),
                "columnCount": len(paste_data[0])
            }
            }
        }
        }]
    }
    service = initService._getService()
    #clear the columns
    service.spreadsheets().batchUpdate(spreadsheetId = spreadsheetId, body=body).execute()
    
    resp = update_tab_with_df(df, sheet_name = sheet_name, spreadsheetId = spreadsheetId, header=header)
    
    return resp

def read_google_sheet(spreadsheetId = None, sheet_name = None):
    """
    This will 
    
    Parameters
    ----------
    spreadsheetId : str, required
        The ID of the spreadsheet to read from
    
    sheet_name : str, optional
        This is the name of the tab/sheet you would like to read from. Without it, it defaults to
        the first sheet in the spreadsheet.
        
    
    Returns
    -------
    Returns a Pandas Dataframe of the sheet with cells formatted as strings.
    """
    if not spreadsheetId:
        raise ValueError('Please specify a spreadsheetId.')

    service = initService._getService()

    #if the sheet name isn't specified, use the first one we find
    if not sheet_name:
        current_state = service.spreadsheets().get(spreadsheetId = spreadsheetId).execute()
        sheet_name = current_state['sheets'][0]['properties']['title']
    
    response = service.spreadsheets().values().get(spreadsheetId = spreadsheetId, range=sheet_name).execute()
    return pd.DataFrame(response['values'][1:len(response['values'])], columns=response['values'][0])

initialize_service._initializeService(initializing=True)