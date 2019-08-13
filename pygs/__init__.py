#!/usr/bin/env python

__author__ = "JP Schultz jp.schultz@gmail.com"
__license__ = "MIT"

#py3 Compatability
try:
    import pygs_tools as pytools
    import initialize_service as init_service
except ImportError:
    from . import pygs_tools as pytools
    from . import initialize_service as init_service



def create_empty_spreadsheet(document_name=None, sheet_name=None, **kwargs):
    """
    This will create an empty Google Sheet and return an object
    containing both the 'key' and the 'url' of the created sheet.

    Parameters
    ----------
    document_name : str, optional
        This will be the name of the Google Sheet. If left blank,
        it will default to 'Untitled Spreadsheet'

    sheet_name : str, optional
        This is the name of the first tab within the Google Sheet Doc.
        If not set, will be Sheet1.


    Returns
    -------
    Returns an object containing both the 'key' for the spreadsheet and the 'url'
    """
    service = init_service.get_service()

    cols = kwargs.pop('cols', 26)
    rows = kwargs.pop('rows', 1000)

    if document_name is None:
        document_name = 'Untitled spreadsheet'

    if sheet_name is None:
        sheet_name = 'Sheet1'

    sheets_info = [{
        'properties': {
            'gridProperties': {
                'columnCount': cols,
                'rowCount': rows
            },
            'index': 0,
            'sheetId': 0,
            'sheetType': 'GRID',
            'title': sheet_name
        }
    }]

    response = service.spreadsheets().create(body={'properties': {'title': document_name},
                                                   'sheets': sheets_info}).execute()
    ret_val = {
        'spreadsheetId': str(response['spreadsheetId']),
        'spreadsheetUrl': str(response['spreadsheetUrl'])
    }

    return ret_val


def get_all_sheet_names(spreadsheetId):
    """
    This will return a list of all of the sheet names.

    Parameters
    ----------

    spreadsheetId : str
        The ID of the Google Sheet you wish to get the Sheet
        names for


    Returns
    -------
    Returns a Python list containing all the sheet names
    """

    sheet_names = pytools.get_all_sheet_names(spreadsheetId)

    return sheet_names


def create_spreadsheet_from_df(df, sheet_name=None, document_name=None, header=True):
    """
    Given a Pandas DataFrame (df), this will create a google sheet
    and name it the document_name and paste
    the DataFrame to the first sheet, which it will name the sheet_name

    Parameters
    ----------
    df : Pandas Dataframe
         This will take the dataframe and output it to a brand new Google Sheet
    sheet_name : str, optional
        The name of the tab within the sheet that you would like the
        DataFrame pasted to. If left blank, will default to 'Sheet1'
    document_name : str, optional
        This will be the name of the Google Sheet. If left blank,
        it will default to 'Untitled Spreadsheet'
    header : bool, optional
        This will determine if the header (column titles) are included
        when pasting data. Defaults to true.
        Setting to False will mean that just the raw data is pasted, starting in the first cell

    Returns
    -------
    Returns an object containing both the 'key' for the spreadsheet and the 'url'

    """
    if df.empty:
        raise ValueError('Please pass in a dataframe with data.')

    # clean/prep the dataframe to go to google sheets
    df = pytools.cleanDF(df)

    if header:
        paste_data = [df.columns.tolist()] + df.values.tolist()
    else:
        paste_data = df.values.tolist()

    total_cells = df.size + len(paste_data[0])

    if total_cells > 2000000:
        raise ValueError(
            'There are more than 2 million cells in this dataframe which cannot be loaded into Google Sheets.')

    if document_name is None:
        document_name = 'Untitled spreadsheet'
    if sheet_name is None:
        sheet_name = 'Sheet1'

    last_col = pytools.getEndCol(paste_data)
    last_row = str(len(paste_data))

    a1notation = sheet_name + '!A1:' + last_col + last_row

    if len(paste_data) * 26 > 2000000:
        cols = len(paste_data[0])
        rows = len(paste_data)
    else:
        cols = 26
        rows = 1000

    service = init_service.get_service()

    new_sheet = create_empty_spreadsheet(document_name=document_name,
                                         sheet_name=sheet_name,
                                         cols=cols,
                                         rows=rows)

    new_sheet_id = new_sheet['spreadsheetId']

    body = {
        "range": a1notation,
        "values": paste_data
    }

    response = service.spreadsheets().values().update(spreadsheetId=new_sheet_id,
                                                      range=a1notation,
                                                      body=body,
                                                      valueInputOption='USER_ENTERED').execute()

    ret_val = {
        'status': 'success',
        'spreadsheetId': str(response['spreadsheetId']),
        'spreadsheetUrl': str(new_sheet['spreadsheetUrl'])
    }

    return ret_val


def update_sheet_with_df(df, sheet_name, spreadsheetId, header=True):
    """
    Given a Pandas DataFrame (df), spreadsheetId and sheet_name, this will
    empty the sheet and paste the dataframe into it.

    Parameters
    ----------
    df : Pandas Dataframe, required
         This will take the dataframe paste it into a defined tab on a Google Sheet
    sheet_name : str, required
        The name of the tab within the sheet that you would like the DataFrame pasted to.
    spreadsheetId : str, required
        The ID of the spreadsheet containing the sheet that you want the DataFrame pasted to.
    header : bool, optional
        This will determine if the header (column titles) are included when pasting data.
        Defaults to true.
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

    # clean/prep the dataframe to go to google sheets
    df = pytools.cleanDF(df)

    if header:
        paste_data = [df.columns.tolist()] + df.values.tolist()
    else:
        paste_data = df.values.tolist()

    total_cells = df.size + len(paste_data[0])

    if total_cells > 2000000:
        raise ValueError('There are more than 2 million cells in \
                        this dataframe which cannot be loaded into Google Sheets.')

    last_col = pytools.getEndCol(paste_data)
    last_row = str(len(paste_data))

    a1notation = sheet_name + '!A1:' + last_col + last_row

    service = init_service.get_service()

    current_state = service.spreadsheets().get(
        spreadsheetId=spreadsheetId).execute()

    found = False
    for sheet in current_state['sheets']:
        if sheet['properties']['title'] == sheet_name:
            current_cols = sheet['properties']['gridProperties']['columnCount']
            current_rows = sheet['properties']['gridProperties']['rowCount']
            sheet_id = sheet['properties']['sheetId']
            found = True
            break

    if not found:
        raise ValueError(
            "Unable to find '{}' in the spreadsheet. Please check the sheet name again.".format(sheet_name))

    # clear the data from the sheet
    clear_range_end_col = pytools.getEndCol([range(current_cols)])

    range_to_clear = "{}!A1:{}{}".format(sheet_name, clear_range_end_col, str(current_rows))

    service.spreadsheets().values().clear(spreadsheetId=spreadsheetId,
                                          range=range_to_clear,
                                          body={}).execute()

    # remove extra columns so we can fit the new DF
    # If the length of the incoming data multiplied by the number of columns there currently are would
    # make the sheet more than 2000000, we need to remove any extra columns, so this deletes those columns
    if (len(paste_data) * current_cols) > 2000000:
        range_dict = {
            'sheetId': sheet_id,
            'dimension': 'COLUMNS',
            'startIndex': 1,
            'endIndex': current_cols - 1
        }
        # request body based off of the dimensions above
        body = {
            "requests": [{
                "deleteDimension": {
                    "range": range_dict
                }
            }]
        }
        # clear the columns
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheetId, body=body).execute()

    # If the incoming data is shorter than the current data, we need to clear the old data first, otherwise
    # we can just paste over it
    if len(paste_data) < current_rows:
        range_dict = {
            'sheetId': sheet_id,
            'dimension': 'ROWS',
            'startIndex': len(paste_data) - 1,
            'endIndex': current_rows
        }
        # request body based off of the dimensions above
        body = {
            "requests": [{
                "deleteDimension": {
                    "range": range_dict
                }
            }]
        }
        # clear the columns
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheetId, body=body).execute()

    body = {
        "range": a1notation,
        "values": paste_data
    }

    response = service.spreadsheets().values().update(
        spreadsheetId=spreadsheetId,
        range=a1notation,
        body=body,
        valueInputOption='USER_ENTERED'
    ).execute()

    ret_val = {
        'status': 'success',
        'spreadsheetId': str(response['spreadsheetId']),
        'spreadsheetUrl': 'https://docs.google.com/spreadsheets/d/' + str(response['spreadsheetId'])
    }

    return ret_val


def create_tab_from_df(df, sheet_name, spreadsheetId, header=True):
    """
    Given a Pandas DataFrame (df), spreadsheetId and sheet name,
    this will create a new tab in the spreadsheet
    with the given data.

    Parameters
    ----------
    df : Pandas Dataframe, required
         This will take the dataframe paste it into a defined tab on a Google Sheet
    sheet_name : str, required
        The name of the tab within the sheet that you want created.
    spreadsheetId : str, required
        The ID of the spreadsheet containing the sheet that you want the DataFrame pasted to.
    header : bool, optional
        This will determine if the header (column titles) are included
        when pasting data. Defaults to true.
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

    # clean/prep the dataframe to go to google sheets
    df = pytools.cleanDF(df)

    if header:
        paste_data = [df.columns.tolist()] + df.values.tolist()
    else:
        paste_data = df.values.tolist()

    total_cells = df.size + len(paste_data[0])

    if total_cells > 2000000:
        raise ValueError('There are more than 2 million cells in \
                                this dataframe which cannot be loaded into Google Sheets.')

    sheet_name = pytools.clean_sheet_name(sheet_name, spreadsheetId)

    body = {
        "requests": [{
            "addSheet": {
                "properties": {
                    "title":
                        sheet_name,
                    "gridProperties": {
                        "rowCount": len(paste_data),
                        "columnCount": len(paste_data[0])
                    }
                }
            }
        }]
    }
    service = init_service.get_service()
    # create the empty sheet
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheetId, body=body).execute()

    resp = update_sheet_with_df(df,
                                sheet_name=sheet_name,
                                spreadsheetId=spreadsheetId,
                                header=header)

    return resp


def read_google_sheet(spreadsheetId=None, sheet_name=None):
    """
    This will read in a Google Sheet to a Pandas DataFrame

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

    service = init_service.get_service()

    # if the sheet name isn't specified, use the first one we find
    if not sheet_name:
        current_state = service.spreadsheets().get(
            spreadsheetId=spreadsheetId).execute()
        sheet_name = current_state['sheets'][0]['properties']['title']

    response = service.spreadsheets().values() \
        .get(spreadsheetId=spreadsheetId, range=sheet_name).execute()
    return pytools.fixResponse(response)


def get_total_cells(spreadsheetId):
    """
    Given a specific spreadsheetId, this will return a count of the number of
    cells in the spreadsheet. Since Google Sheets has a limit, this can
    be used to determine if you will go over it by adding more.

    Parameters
    ----------
    spreadsheetId : str, required
        The ID of the spreadsheet to read from


    Returns
    -------
    Returns int representing the number of cells in the sheet.
    """
    if not spreadsheetId:
        raise ValueError('Please specify a spreadsheetId.')

    service = init_service.get_service()
    sheet_info = service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()

    total_cells = 0
    for sheet in sheet_info['sheets']:
        columns = sheet['properties']['gridProperties']['columnCount']
        rows = sheet['properties']['gridProperties']['rowCount']
        total_cells += columns * rows

    return total_cells


init_service.initialize_service(initializing=True)
