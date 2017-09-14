import pandas as pd
import xlsxwriter
import os
from datetime import datetime

# find the files neede to process
ok_file = ""
cancel_file = ""
booking_file = ""
files = [f for f in os.listdir('.') if os.path.isfile(f)]
for f in files:
    if " ok " in f and f[:2] != "~$":
        ok_file = f
    if " cancel " in f and f[:2] != "~$":
        cancel_file = f
    if "Booking.com" in f and f[:2] != "~$":
        booking_file = f

msg = "Files detected:\n" + booking_file + "\n" + ok_file + "\n" + cancel_file + "\n" + "Press Enter key to begin the process."
a = raw_input(msg)


# developed with python 3.5 for pyinstaller function
#############################################################
# check conditions and print out
def same_name(booking_name, hotel_name):
    if hotel_name[:5] == 'name\n':
        hotel_name = hotel_name[5:]

    hotel_name = hotel_name.split(",")
    hotel_name = hotel_name[1][1:] + " " + hotel_name[0]

    if hotel_name == booking_name:
        # print("h:", hotel_name, "b:", booking_name)
        return True
    else:
        return False


def calc_hotel_price(hotel_price, arrival, departure):
    # print("input: ", booking_price, hotel_price, arrival, departure)

    if departure[:2] == "De":
        departure = departure[9:]
        arrival = arrival[7:]
        hotel_price = hotel_price[4:]

    date_format = "%m/%d/%Y"
    a = datetime.strptime(departure, date_format)
    b = datetime.strptime(arrival, date_format)
    delta = a - b
    # print("hotel price: ", float(delta.days) * float(hotel_price), "booking price: ", float(booking_price))
    # print(float(delta.days) * float(hotel_price) == float(booking_price))
    return (float(delta.days) * float(hotel_price))


######################################################################################
match = 0
diff_price = 0
canceled = 0
not_found = 0

# load and prepare the booking df
xl = pd.ExcelFile(booking_file)
booking_df = xl.parse("Sheet1", header=0, keep_default_na=False)

booking_col_names = list(booking_df)
print(booking_col_names)
for name in booking_col_names:
    a = name.split(" ")
    if a[-1] == "":
        b = " ".join(a[:-1])
    else:
        b = " ".join(a)
    booking_df = booking_df.rename(columns={name: b})

booking_col_names = list(booking_df)
print(booking_col_names)


b_df = booking_df[booking_df['Status'] == 'ok']
print('booking df created')


guest_col_name =""
if "Guest name(s)" in list(b_df):
    guest_col_name = 'Guest name(s)'
    print("YESSSSSSSSSSSSSS")
else:
    guest_col_name = 'Guest Name(s)'
    print("22222222222222")

for index, row in b_df.iterrows():
    if row[guest_col_name] == '':
        name = str(row['Booked by'])
        guest_name = name.split(",")
        b_df.loc[index, guest_col_name] = guest_name[1][1:] + " " + guest_name[0]

##########################################
# loading the ok DF
xl = pd.ExcelFile(ok_file)
ok_df = xl.parse("Sheet1", header=2, keep_default_na=False, )
print(list(ok_df))
o_df = ok_df.loc[(ok_df['Name'] != "") & (ok_df['Name'] != "Name") &
                 (ok_df['Name'] != "Guest List Detail") &
                 (ok_df['Name'] != "30031 - DAYS INN PHILADELPHIA.") &
                 (ok_df['Name'] != "Name Arrival Departure Rate")]

# combine ok name
for row in range(1, len(o_df)):
    if o_df.iloc[row]['Rate'] == '':
        o_df.iloc[row - 1]['Name'] = o_df.iloc[row - 1]['Name'] + " " + o_df.iloc[row]['Name']

# clear all empty rows and finalize df
o_df = o_df.loc[(o_df['Rate'] != "")]

# print o_df
# for index, row in o_df.iterrows():
#     print("index:::", index, row['Name'], row['Arrival'], row['Departure'], row['Rate'])
#

##########################
# loading the cancel DF
xl = pd.ExcelFile(cancel_file)
c_df = xl.parse("Sheet1", header=2, keep_default_na=False, )
print(list(c_df))
c_df = c_df.loc[(c_df['Name'] != "") & (c_df['Name'] != "Name") &
                (c_df['Name'] != "Guest List Detail") &
                (c_df['Name'] != "30031 - DAYS INN PHILADELPHIA.") &
                (c_df['Name'] != "Name Arrival Departure Rate")]

# combine ok name
for row in range(1, len(c_df)):
    if c_df.iloc[row]['Rate'] == '':
        c_df.iloc[row - 1]['Name'] = c_df.iloc[row - 1]['Name'] + " " + c_df.iloc[row]['Name']

# clear all empty rows and finalize df
c_df = c_df.loc[(c_df['Rate'] != "")]

# print o_df
# for index, row in c_df.iterrows():
#     print("index:::", index, row['Name'], row['Arrival'], row['Departure'], row['Rate'])

# ok_file = ""
# cancel_file = ""
# booking_file = ""
# files = [f for f in os.listdir('.') if os.path.isfile(f)]
# for f in files:
#     if " ok " in f and f[:2] != "~$":
#         ok_file = f
#     if " cancel " in f and f[:2] != "~$":
#         cancel_file = f
#     if " booking " in f and f[:2] != "~$":
#         booking_file = f

#########################################
book_col_name = list(b_df)[0]

print("begin comparison")
# create worksheet
workbook = xlsxwriter.Workbook('Expenses.xlsx')
worksheet = workbook.add_worksheet()
rowm = 0
coln = 0
worksheet.write(rowm, coln, "Confirmation Number")
worksheet.write(rowm, coln + 1, "CRS Number")
worksheet.write(rowm, coln + 2, "Name")
worksheet.write(rowm, coln + 3, "Booking.com Price")
worksheet.write(rowm, coln + 4, "Description")
rowm += 1

workbook_nf = xlsxwriter.Workbook('Customers Not found.xlsx')
worksheet_nf = workbook_nf.add_worksheet()
row_nf = 0
total_price_difference = 0.00
# # print b_df
# print(list(b_df))
# for index, row in b_df.iterrows():
#     print("index:::", index, row['Guest name(s)'].lower(),row['Check-in'],row['Check-out'],row['Price'])
#

# check loop
for index, row in b_df.iterrows():

    # initialzie what to compare
    name = row[guest_col_name]
    check_in_date = row['Check-in']
    check_out_date = row['Check-out']
    # print("check_in_date: ", check_in_date)
    check_in_date = check_in_date.split("-")
    # print("check_in_date after split: ", check_in_date)
    check_out_date = check_out_date.split("-")
    price = row['Price'][:-3]
    found = False

    # check if in ok, but different price
    for index_ok, row_ok in o_df.iterrows():
        if (same_name(name.lower(), row_ok['Name'].lower())):

            if row_ok['Departure'][:2] == "De":
                departure = row_ok['Departure'][9:]
                arrival = row_ok['Arrival'][7:]
            else:
                departure = row_ok['Departure']
                arrival = row_ok['Arrival']
            # hotel date
            date_format = "%m/%d/%Y"
            h_departure = datetime.strptime(departure, date_format)
            h_arrival = datetime.strptime(arrival, date_format)

            # booking.com date

            b_checkin = datetime(int(check_in_date[0]), int(check_in_date[1]), int(check_in_date[2]))
            b_checkout = datetime(int(check_out_date[0]), int(check_out_date[1]), int(check_out_date[2]))
            found = True
            match += 1

            if not ((h_arrival == b_checkin) & (h_departure == b_checkout)):
                worksheet.write(rowm, coln, row[book_col_name])
                worksheet.write(rowm, coln + 1, row_ok['CRS#'])
                worksheet.write(rowm, coln + 2, name)
                worksheet.write(rowm, coln + 3, row['Price'])
                worksheet.write(rowm, coln + 4, "Checked in, but different date")
                print(name, " Found in OK file, but different date")
                rowm += 1
                break

    if (found == False):
        # check if in canceled
        for index_c, row_c in c_df.iterrows():
            if (same_name(name.lower(), row_c['Name'].lower())):
                # print(name + "**BAD***********************************************")
                # worksheet.write(rowm, coln, name)
                #
                # worksheet.write(rowm, coln + 4, "Found in Cancel file.")
                # rowm += 1
                # found = True
                # canceled += 1

                if row_c['Departure'][:2] == "De":
                    departure = row_c['Departure'][9:]
                    arrival = row_c['Arrival'][7:]
                else:
                    departure = row_c['Departure']
                    arrival = row_c['Arrival']
                    # hotel date
                date_format = "%m/%d/%Y"
                h_departure = datetime.strptime(departure, date_format)
                h_arrival = datetime.strptime(arrival, date_format)

                # booking.com date

                b_checkin = datetime(int(check_in_date[0]), int(check_in_date[1]), int(check_in_date[2]))
                b_checkout = datetime(int(check_out_date[0]), int(check_out_date[1]), int(check_out_date[2]))

                if (h_arrival == b_checkin) & (h_departure == b_checkout):
                    found = True

                    print(name, price, "found in CANCEL file")
                    worksheet.write(rowm, coln, row[book_col_name])
                    worksheet.write(rowm, coln + 1, row_c['CRS#'])
                    worksheet.write(rowm, coln + 2, name)
                    worksheet.write(rowm, coln + 3, row['Price'])
                    worksheet.write(rowm, coln + 4, "Canceled")
                    print(name, " Found in cancel file")
                    rowm += 1
                    break

    if found == False:
        print(name, " Can't find customer Name")
        not_found += 1
        worksheet_nf.write(row_nf, 0, name)
        row_nf += 1

workbook.close()
workbook_nf.close()

print("match: ", match)
print("diff_price: ", diff_price)
print("canceled: ", canceled)
print("not found: ", not_found)
