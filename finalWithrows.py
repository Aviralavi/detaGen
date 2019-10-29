# program to generate random account numbers
import xlrd
from xlwt import Workbook


# read excel fileexpected_Count_For_Single_Row
loc = "C:\\Users\\aviverma\\Desktop\\dataGEN\DATA.xlsx"

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_name('xyz')

wrkb = Workbook()
sheet1 = wrkb.add_sheet('sheet1')

sheet1.write(0, 0, 'CURRENCY')
sheet1.write(0, 1, 'SUB_PRODID')
sheet1.write(0, 2, 'CUSTOMER')
sheet1.write(0, 3, 'ACCOUNT NUMBER')
def item_Handler(item):
    new_List = item.split('-')
    # print(new_List)
    return new_List


def getList(input_column):
    # Extracting all columns name
    get_list = []
    notify = False
    for i in range(sheet.ncols):
        # print(sheet.cell_value(0, i))
        if notify is True:
            return get_list
            exit()
        if sheet.cell_value(0, i) == input_column:
            # Extract first column
            for j in range(sheet.nrows):
                get_list.append(sheet.cell_value(j, i))
                # print(get_list)
                notify = True


# --------------------------------------------------------
def currencyList():
    global get_currency_list_modified
    get_currency_list_modified = []
    # print("currency count")
    get_currency_ist = getList("CURRENCY")
    lengthpro = len(get_currency_ist)
    for typeindex in range(lengthpro):
        if type(get_currency_ist[typeindex]) == float:
            get_currency_ist[typeindex] = int(get_currency_ist[typeindex])

    for index in get_currency_ist:
        if index != '':
            get_currency_list_modified.append(index)
    # print(get_currency_list_modified)


def subproductList():
    global get_subproduct_list_modified
    get_subproduct_list_modified = []
    # print("currency count")
    get_subproduct_list = getList("SUB_PRODID")
    lengthpro = len(get_subproduct_list)
    for typeindex in range(lengthpro):
        if type(get_subproduct_list[typeindex]) == float:
            get_subproduct_list[typeindex] = int(get_subproduct_list[typeindex])

    for index in get_subproduct_list:
        if index != '':
            get_subproduct_list_modified.append(index)
    # print(get_subproduct_list_modified)


def customerList():
    global get_customer_list_modified
    get_customer_list_modified = []
    # print("customer count")
    get_customer_list = getList("CUSTOMER")
    lengthpro = len(get_customer_list)
    for typeindex in range(lengthpro):
        if type(get_customer_list[typeindex]) == float:
            get_customer_list[typeindex] = int(get_customer_list[typeindex])

    for index in get_customer_list:
        if index != '':
            get_customer_list_modified.append(index)
    # print(get_customer_list_modified)

# --------------------------------------------------------

# sheet value
value = sheet.cell_value(0, 0)

# sheet row number
column_count = sheet.ncols
row_count = str(sheet.nrows - 1)
# print(column_count)
# print(row_count)

# Extracting all columns name
# for i in range(sheet.ncols):
    # print(sheet.cell_value(0, i))

# Extract first column
# for i in range(sheet.nrows):
    # print(sheet.cell_value(i, 0))

# --------------------------------------------------------------
# dummy input
# rowid_range = str(input("input row range"))
total_account_count = int(input("Enter total number of account to be generated: "))
if total_account_count <0:
    print("enter positive value")
    exit()
rowid_range = ''

if rowid_range is '':
    rowid_range = str('1' + '-' + row_count)
    # print(rowid_range)

total_count = 0
my_list = rowid_range.split(',')


def get_input_Row_Count():
    global total_count
    total_count += 1


try:
    for item in my_list:
        if '-' in item != -1:
            new_List = item_Handler(item)
            new_List = list(range(int(new_List[0]), int(new_List[1]) + 1))
            # EXECUTE ROW
            for index in new_List:
                get_input_Row_Count()
        else:
            # EXECUTE ROW
            get_input_Row_Count()
except Exception as error:
    print(error)


# --------------------------------------------------------------

def get_Row_Value(row_value):
    row_value = int(row_value)
    # print(sheet.row_values(row_value))


# Extract a particular row value
# print(sheet.row_values(1))
# print(sheet.row_values(0))

# rowid concept input extract
my_list = rowid_range.split(',')
# print(my_list)
try:
    for item in my_list:
        if '-' in item != -1:
            new_List = item_Handler(item)
            new_List = list(range(int(new_List[0]), int(new_List[1]) + 1))
            # EXECUTE ROW
            for index in new_List:
                get_Row_Value(index)
        else:
            # EXECUTE ROW
            get_Row_Value(item)
except Exception as error:
    print(error)


def expected_Count_For_Single_Row(total_account_count, total_customer_count):
    count_For_Single_Row = float(total_account_count) / float(total_customer_count)
    single_row_add = int(count_For_Single_Row)
    if count_For_Single_Row > single_row_add:
        count_For_Single_Row = count_For_Single_Row + 1
    count_For_Single_Row = int(count_For_Single_Row)
    return count_For_Single_Row

currencyList()
customerList()
subproductList()

total_customer_count = len(get_customer_list_modified)-1
# print(total_customer_count)
count_For_Single_Row = expected_Count_For_Single_Row(total_account_count, total_customer_count)
# print(count_For_Single_Row)



def currency_check(get_currency_list_modified_pass):
    # for index in get_currency_list_modified:
    if get_currency_list_modified_pass == "USD":
        return '01'
    if get_currency_list_modified_pass == "EUR":
        return '02'
    if get_currency_list_modified_pass == "JPY":
        return '03'
    if get_currency_list_modified_pass == "GBP":
        return '04'
    if get_currency_list_modified_pass == "CAD":
        return '05'
        exit()


def accountGenerator(get_currency_list_modified_pass, get_subproduct_list_modified_pass,
                     get_customer_list_modified_pass, loop_count):
    global account_number
    global get_sub
    global get_cust
    # account_number = []
    if len(loop_count) == 1:
        loop_count = '0' + loop_count

    currency_number = currency_check(get_currency_list_modified_pass)
    # print(get_subproduct_list_modified_pass)
    sub_length = len(get_subproduct_list_modified_pass)

    if sub_length < 4:
        for index in range(4-sub_length):
            get_sub = get_subproduct_list_modified_pass + '0'
    else:
        get_sub = get_subproduct_list_modified_pass

    cust_length = len(get_customer_list_modified_pass)

    if cust_length < 7:
        for index in range(7-cust_length):
            get_customer_list_modified_pass = get_customer_list_modified_pass + '0'
            get_cust = get_customer_list_modified_pass
    else:
        get_cust = get_customer_list_modified_pass
    # print(loop_count)
    modified_subproduct = get_sub[0] + get_sub[1] + get_sub[2] + get_sub[3]
    modified_customer = get_cust[0] + get_cust[1] + get_cust[2] + get_cust[3] + get_cust[4] + get_cust[5] + get_cust[6]
    if int(loop_count) > 99:
        modified_customer = modified_customer[0:6]
        if int(loop_count) > 999:
            modified_customer = modified_customer[0:5]
            if int(loop_count) > 9999:
                modified_customer = modified_customer[0:4]

    return currency_number + modified_subproduct + modified_customer + loop_count

global total_loop_count
total_loop_count = 0
n=0
try:
    for i in range(len(get_customer_list_modified) - 1):
        if (total_account_count == total_loop_count):
            break
        store_customer = get_customer_list_modified[i + 1]
        loop_count = -1
        global inner_loop_count
        inner_loop_count = 0
        while (inner_loop_count < count_For_Single_Row):
            if total_account_count == total_loop_count:
                break
            loop_count += 1
            for j in range(len(get_currency_list_modified) - 1):
                if (inner_loop_count == count_For_Single_Row or total_account_count == total_loop_count):
                    break
                store_currency = get_currency_list_modified[j + 1]
                for k in range(len(get_subproduct_list_modified) - 1):
                    if (inner_loop_count == count_For_Single_Row or total_account_count == total_loop_count):
                        break
                    store_subproduct = get_subproduct_list_modified[k + 1]
                    ret_value = accountGenerator(str(store_currency), str(store_subproduct), str(store_customer), str(loop_count))
                    print(ret_value)
                    sheet1.write(total_loop_count+1, 0, str(store_currency))
                    sheet1.write(total_loop_count+1, 1, str(store_subproduct))
                    sheet1.write(total_loop_count+1, 2, str(store_customer))
                    sheet1.write(total_loop_count+1, 3, str(ret_value))
                    inner_loop_count += 1
                    total_loop_count += 1
                    if (inner_loop_count == count_For_Single_Row or total_account_count == total_loop_count):
                        break
except Exception as e:
    print(e)

wrkb.save("C:\\Users\\aviverma\\Desktop\\dataGEN\generated_accounts.xls")

