import pandas as pd
import xlsxwriter
import os
import datetime

# find the files neede to process
hotel_file = ""
# cancel_file = ""
booking_file = ""
files = [f for f in os.listdir('.') if os.path.isfile(f)]

for f in files:
    f = f.lower()
    if "report" in f and f[:2] != "~$":
        hotel_file = f
    if "expedia" in f and f[:2] != "~$":
        booking_file = f

msg = "Files detected:\n" + booking_file + "\n" + hotel_file + "\n" + "Press Enter key to begin the process."
a = raw_input(msg)


# developed with python 3.5 for pyinstaller function
#############################################################
# check conditions and print out
def same_name(booking_name, hotel_name):
    if hotel_name[:5] == 'name\n':
        hotel_name = hotel_name[5:]

    hotel_name = hotel_name.split(",")
    hotel_name = hotel_name[1][:] + " " + hotel_name[0]

    # print("H:", hotel_name, " W:", booking_name)
    if hotel_name == booking_name:
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
    a = datetime.datetime.strptime(departure, date_format)
    b = datetime.datetime.strptime(arrival, date_format)
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
booking_df = xl.parse(header=0, keep_default_na=False)

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
# print(booking_col_names)

website_df = booking_df[booking_df['Status'] == 'Booked']
#
# print("expedia_file")
# for index, row in website_df.iterrows():
#     print("index:::", index, row['Guest'], row['Check-In'], row['Check-Out'])

##########################################
# loading the ok DF
xl = pd.ExcelFile(hotel_file)
o_df = xl.parse(header=0, keep_default_na=False, )

#
# # print o_df
# print("hotel_file")
# for index, row in o_df.iterrows():
#     print("index:::", index, row['GuestName'], row['ArrivalDt'], row['DaysStay'], row['CancelDt'])

# #########################################
print("begin comparison")
# create worksheet
workbook = xlsxwriter.Workbook('Expenses.xlsx')
worksheet = workbook.add_worksheet()
rowm = 0
coln = 0
worksheet.write(rowm, coln, "Confirmation Number")
worksheet.write(rowm, coln + 1, "CRS Number")
worksheet.write(rowm, coln + 2, "Name")
worksheet.write(rowm, coln + 3, "Description")

rowm += 1

workbook_nf = xlsxwriter.Workbook('Customers Not found.xlsx')
worksheet_nf = workbook_nf.add_worksheet()
row_nf = 0

print("before check loop")
good = 0
##########################################################################
for index, row in website_df.iterrows():

    # initialzie what to compare
    name = row['Guest']
    check_in_date = row['Check-In']
    check_out_date = row['Check-Out']
    # print("check_in_date: ", check_in_date)
    check_in_date = str(check_in_date).split("-")
    # print(check_in_date)
    # print("check_in_date after split: ", check_in_date)
    check_out_date = str(check_out_date).split("-")
    # print(check_out_date)
    # price = row['Price'][:-3]
    found = False

    # check if in ok, but different price
    for index_ok, row_ok in o_df.iterrows():
        if (same_name(name.lower(), row_ok['GuestName'].lower())):
            found = True

            # if name matched
            # check if in canceled
            print("checking if in cancel file first")
            if row_ok['CancelDt'] != "":
                print("ROOM canceled", row_ok['CancelDt'])

                worksheet.write(rowm, coln, row['Confirmation #'])
                worksheet.write(rowm, coln + 1, row_ok['CRSBookNum'])
                worksheet.write(rowm, coln + 2, name)
                worksheet.write(rowm, coln + 3, "Cancelled")
                print(name, " found in CANCEL file")
                rowm += 1
                canceled += 1
                break

            # if not found in cancel file, we check the date
            # hotel date
            date_format = "%m/%d/%Y"
            arrival = row_ok['ArrivalDt']
            h_arrival = datetime.datetime.strptime(arrival, date_format)
            h_departure = h_arrival + datetime.timedelta(days=row_ok['DaysStay'])
            # + int(row_ok['DaysStay'])

            # booking.com date

            b_checkin = datetime.datetime(int(check_in_date[0]), int(check_in_date[1]), int(check_in_date[2][:2]))
            b_checkout = datetime.datetime(int(check_out_date[0]), int(check_out_date[1]), int(check_out_date[2][:2]))
            found = True
            match += 1

            print('h_arrival: ', h_arrival)
            print('h_departure: ', h_departure)
            print("b_checkin: ", b_checkin)
            print("b_checkout", b_checkout)
            if not ((h_arrival == b_checkin) & (h_departure == b_checkout)):
                worksheet.write(rowm, coln, row['Confirmation #'])
                worksheet.write(rowm, coln + 1, row_ok['CRSBookNum'])
                worksheet.write(rowm, coln + 2, name)
                worksheet.write(rowm, coln + 3, "Checked in, but different date")
                print(name, " Found in OK file, but different date")
                rowm += 1
                break
            else:
                good += 1
                break

    if found == False:
        # print(name, " Can't find customer Name")
        not_found += 1
        worksheet_nf.write(row_nf, 0, name)
        row_nf += 1
workbook.close()
workbook_nf.close()
print("good:", good)
print("match: ", match)
print("canceled: ", canceled)
print("not found: ", not_found)
