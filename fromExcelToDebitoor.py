import os
import xlrd
from ordereddict import OrderedDict
import simplejson as json
 
# Open the workbook and select the first worksheet
wb = xlrd.open_workbook('') #WRITE YOUR excel filename here
sh = wb.sheet_by_index(0)
 
#WRITE YOUR access Token below
accessToken = ""

# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, sh.nrows):
    product = OrderedDict()
    row_values = sh.row_values(rownum)
    product['sku'] = row_values[0]
    product['name'] = row_values[1]
    product['description'] = row_values[2]
    product['netUnitSalesPrice'] = str(round(float(row_values[3]), 3))
    product['grossUnitSalesPrice'] = str(round(float(row_values[4]), 3))
    product['netUnitCostPrice'] = str(round(float(row_values[5]), 3))
    product['rate'] = row_values[6]
    product['taxEnabled'] = True

    j = json.dumps(product) #Convert to json
    print j
    print "Making petition..."
    command = "curl -X POST https://api.debitoor.com/api/v1.0/products/?access_token=" + acessToken + " -d "  "\'" + j +  "\'"
    os.system(command)
    print "\n\n\n"