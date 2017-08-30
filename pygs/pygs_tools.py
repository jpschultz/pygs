#!/usr/bin/env python
import string
import initialize_service as initService
def _cleanSheetName(sheet_name, spreadsheetId):
    service = initService._getService()
    current_state = service.spreadsheets().get(spreadsheetId = spreadsheetId).execute()
    tot_names = []
    for sheet in current_state['sheets']:
        if sheet_name.lower() == sheet['properties']['title'].lower():
            try:
                name_split = sheet['properties']['title'].split(sheet_name + "_")
                if len(name_split) > 1:
                    tot_names.append(int(name_split[-1])) 
                else:
                    tot_names.append(0) 
            except ValueError:
                pass
    if len(tot_names) == 0:
        return sheet_name
    new_sheet_name = sheet_name + "_" + str(max(tot_names)+1)
    return new_sheet_name

def _getEndCol(two_dim_array):
    array_length = len(two_dim_array[0])
    
    if array_length > 702:
        raise ValueError("You have more than 702 columns which would be past column 'ZZ'. Please make smarter choices.")
    
    lookup_dict = {}
    for pos, char in enumerate(string.ascii_uppercase):
        lookup_dict[pos+1] = char
        
    if array_length <= 26:
        return lookup_dict[array_length]
    
    if array_length % 26 == 0:
        first_letter = lookup_dict[(array_length%26)+(array_length/26)-1]
        second_letter = 'Z'
    else:
        first_letter = lookup_dict[array_length/26]
        second_letter = lookup_dict[array_length%26]
    
    return first_letter + second_letter

def _cleanDF(df):
    #fill any NA's
    df = df.fillna('')
    #try to encode it as a string, if all else fails, do utf-8
    #TODO - check how utf-8 handles datetimes
    try:
        df = df.astype(str)
    except UnicodeEncodeError:
        for col in df.columns:
            df[col] = df[col].apply(lambda x: x.encode('utf-8'))
    return df