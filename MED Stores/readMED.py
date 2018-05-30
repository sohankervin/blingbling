#! /usr/bin/python
"""
sophie hannah
readMED.py - this will read an excel doc containing different types of marijuana
businesses in Colorado and will output their business type and info, which will then be utilized in
another .py program
"""

import openpyxl, pprint, errno


print 'Opening MED Stores workbook...'
try :

    workbook = openpyxl.load_workbook('/home/sophie/Stores 0501208.xlsx')
    worksheet = workbook.get_sheet_by_name('Sheet1')
    storeData = {}
    # fill in storeData dict with a store's name and DBA
    print 'Reading rows...'
    for row in range(2, worksheet.max_row + 1) :
        # fill in row data 
        licensee = worksheet['A' + str(row)].value
        dba = worksheet['B' + str(row)].value
        license_num = worksheet['C' + str(row)].value
 
        """ struct of dict : {'DBA' : {'licensee': 'name', 'license number' : license_num}    }
        >>> storeData['DBA_name']['licensee_name'] returns licensee name
        >>> storeData['DBA_name]['license num'] returns license num"""
        # initialize the dictionary keys and values 
        storeData.setdefault(dba, {'licensee' : licensee, 'license number' : str(license_num) })

    """ create a text file to dump DBA, license number, and licensee so we can find out data associated with 
    with DBA, and return it as a string"""
    print'Writing out results to an executatable .py file'
    medFile = open('med2018stores.py', 'w')
    medFile.write('affiliatedData = ' + pprint.pformat(storeData))

    medFile.close()
    print'Done'

except IOError :
    print "the file does not exist, plesae load it first!"

