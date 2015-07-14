'''
Created on Jun 2, 2015

@author: grovesr
'''

if __name__ == '__main__':
    pass

import re, xlrd, pytz
from datetime import datetime

class XlrdutilsError(Exception):
    pass

class XlrdutilsDateParseError(XlrdutilsError): pass
    
class XlrdutilsReadHeaderError(XlrdutilsError): pass
    
class XlrdutilsOpenWorkbookError(XlrdutilsError): pass
    
class XlrdutilsOpenSheetError(XlrdutilsError): pass

class XlrdutilsInvalidInputsError(XlrdutilsError): pass

def open_workbook(filename=None, file_contents=None):
    try:
        if filename:
            workbook=xlrd.open_workbook(filename=filename)
        elif file_contents:
            workbook=xlrd.open_workbook(file_contents=file_contents)
        else:
            raise XlrdutilsInvalidInputsError('failed to open the workbook. No filename or file_contents provided.')
    except xlrd.XLRDError as e:
        raise XlrdutilsOpenWorkbookError('failed to open the workbook. %s' % repr(e))
    return workbook

def read_header(sheet,headerKeys=[]):

    foundHeaders=0
    for rowIndex in range(sheet.nrows):
        thisLineHeaders=[]
        for colIndex in range(sheet.ncols):
            cell=sheet.cell(rowIndex,colIndex).value
            if foundHeaders < len(headerKeys):
                for key in headerKeys:
                    if re.match(key,str(cell),re.IGNORECASE):
                        # we found a header cell
                        foundHeaders += 1
                        break
            thisLineHeaders.append(cell)
        if foundHeaders == len(headerKeys):
            return(thisLineHeaders,rowIndex)
    # No headers in the file
    raise XlrdutilsReadHeaderError(
        'unable to find a valid header row in the spreadsheet.  Looking for a row that contains at least: %s' 
        % repr(headerKeys))

def read_lines(workbook, sheet='', headerKeys=[], zone=None):
    # 
    try:
        sheet = workbook.sheet_by_name(sheet)
    except xlrd.XLRDError:
        if workbook.nsheets == 1:
            try:
                sheet = workbook.sheet_by_index(0)
            except XlrdutilsOpenSheetError:
                raise XlrdutilsOpenSheetError('failed to open the sheet named %s in the workbook because it doesn''t exist.' % sheet)
                #return -1
        else:
            raise XlrdutilsOpenSheetError('failed to open the sheet named %s in a workbook with multiple sheets' % sheet)
            #return -1
    headers,headerLine=read_header(sheet,headerKeys)
    data={}
    #fill dict with empty lists, one for each header key
    for header in headers:
        data[header]=[]
    for rowIndex in range(sheet.nrows):
        # Find the header row
        if rowIndex > headerLine:
            # Assume rows after the header row contain line items
            # run through the columns and add the data to the data dict 
            for colIndex in range(sheet.ncols):
                cell=sheet.cell(rowIndex,colIndex)
                # parse the cell information base on cell type
                if cell.ctype == xlrd.XL_CELL_TEXT:
                    data[headers[colIndex]].append(cell.value.strip())
                elif cell.ctype == xlrd.XL_CELL_EMPTY:
                    data[headers[colIndex]].append('')
                elif cell.ctype == xlrd.XL_CELL_NUMBER:
                    data[headers[colIndex]].append(cell.value)
                elif cell.ctype == xlrd.XL_CELL_DATE:
                    data[headers[colIndex]].append(parse_date(workbook,
                                                              cell.value,
                                                              row=rowIndex,
                                                              col=colIndex,
                                                              zone=zone))
                else:
                    # unspecified cell type, just output a blank
                    data[headers[colIndex]].append('')
    return data 

def parse_date(workbook,cell,row=0,col=0,zone=None):
    #format: excel date object
    if not zone:
        zone='UTC'
    localZone=pytz.timezone(zone)
    if isinstance(cell,str) or isinstance(cell,unicode):
        if len(cell) == 0:
            # just return the empty string
            return cell
        else:
            # if this is a string make sure we can parse it into a date
            try:
                dateVal=datetime.strptime(cell,'%m/%d/%y %H:%M:%S')
            except Exception as e:
                raise XlrdutilsDateParseError('failed to parse a string date into a timezone object at cell(%d,%d). %s' % (row, col,repr(e)))
            dateTuple=dateVal.year,dateVal.month,dateVal.day,dateVal.hour,\
            dateVal.minute,dateVal.second
            dateVal=datetime(*dateTuple)
            # tag the date as being from the local time zone
            dateVal=localZone.localize(dateVal, is_dst=0)
            # convert it to UTC time zone for saving to database
            dateVal=dateVal.astimezone(pytz.utc)
            return dateVal
    else:
        try:
            timeVal =xlrd.xldate_as_tuple(cell,workbook.datemode)
        except Exception as e:
            raise XlrdutilsDateParseError('failed to parse a string date into a timezone object at cell(%d,%d). %s' % (row, col,repr(e)))
        dateVal=datetime(*timeVal)
        # tag the date as being from the local time zone
        dateVal=localZone.localize(dateVal, is_dst=0)
        # convert it to UTC time zone for saving to database
        dateVal=dateVal.astimezone(pytz.utc)
    return dateVal
    
