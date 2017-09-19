import pandas as pd
import xlsxwriter
import os
import datetime

# find the files neede to process
hotel_file = ""
website_file = ""
files = [f for f in os.listdir('.') if os.path.isfile(f)]
for f in files:
    if "report" in f and f[:2] != "~$":
        hotel_file = f
    if "Booking.com" in f and f[:2] != "~$":
        website_file = f

msg = "Files detected:\n" + website_file + "\n" + hotel_file + "\n" + "\n" + "Press Enter key to begin the process."
a = raw_input(msg)


# developed with python 3.5 for pyinstaller function
#############################################################
# check conditions and print out
def same_name(booking_name, hotel_name):
    if hotel_name[:5] == 'name\n':
        hotel_name = hotel_name[5:]

    hotel_name = hotel_name.split(",")
    hotel_name = hotel_name[1][:] + " " + hotel_name[0]

    if hotel_name == booking_name:
        # print("h:", hotel_name, "b:", booking_name)
        return True
    else:
        return False


######################################################################################
match = 0
canceled = 0
not_found = 0

# load and prepare the booking df
xl = pd.ExcelFile(website_file)
website_df = xl.parse(header=0, keep_default_na=False)

website_col_names = list(website_df)
print(website_col_names)
for name in website_col_names:
    a = name.split(" ")
    if a[-1] == "":
        b = " ".join(a[:-1])
    else:
        b = " ".join(a)
    website_df = website_df.rename(columns={name: b})

website_col_names = list(website_df)
print(website_col_names)

web_df = website_df[website_df['Status'] == 'ok']
print('booking df created')

guest_col_name = ""
if "Guest name(s)" in list(web_df):
    guest_col_name = 'Guest name(s)'
    print("YESSSSSSSSSSSSSS")
else:
    guest_col_name = 'Guest Name(s)'
    print("22222222222222")

for index, row in web_df.iterrows():
    if row[guest_col_name] == '':
        name = row['Booked by']
        # .encode('utf-8')
        guest_name = name.split(",")
        web_df.loc[index, guest_col_name] = guest_name[1][1:] + " " + guest_name[0]

##########################################
# loading the ok DF
xl = pd.ExcelFile(hotel_file)
o_df = xl.parse(header=0, keep_default_na=False, )


#########################################
book_col_name = list(web_df)[0]

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

# # print web_df
# print(list(web_df))
# for index, row in web_df.iterrows():
#     print("index:::", index, row['Guest name(s)'].lower(),row['Check-in'],row['Check-out'],row['Price'])
#

# check loop
good = 0
for index, row in web_df.iterrows():

    # initialzie what to compare
    name = row[guest_col_name]
    check_in_date = row['Check-in']
    check_out_date = row['Check-out']
    # print("check_in_date: ", check_in_date)
    check_in_date = check_in_date.split("-")
    # print("check_in_date after split: ", check_in_date)
    check_out_date = check_out_date.split("-")
    found = False

    # check if in ok, but different price
    for index_ok, row_ok in o_df.iterrows():
        if (same_name(name.lower(), row_ok['GuestName'].lower())):
            found = True

            if row_ok['CancelDt'] != "":
                print("ROOM canceled", row_ok['CancelDt'])

                worksheet.write(rowm, coln, row['Book number'])
                worksheet.write(rowm, coln + 1, row_ok['CRSBookNum'])
                worksheet.write(rowm, coln + 2, name)
                worksheet.write(rowm, coln + 3, row['Price'])
                worksheet.write(rowm, coln + 4, "Cancelled")

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
            b_checkout = datetime.datetime(int(check_out_date[0]), int(check_out_date[1]),
                                           int(check_out_date[2][:2]))
            found = True
            match += 1

            # print('h_arrival: ', h_arrival)
            # print('h_departure: ', h_departure)
            # print("b_checkin: ", b_checkin)
            # print("b_checkout", b_checkout)

            if not ((h_arrival == b_checkin) & (h_departure == b_checkout)):
                worksheet.write(rowm, coln, row['Book number'])
                worksheet.write(rowm, coln + 1, row_ok['CRSBookNum'])
                worksheet.write(rowm, coln + 2, name)
                worksheet.write(rowm, coln + 3, row['Price'])
                worksheet.write(rowm, coln + 4, "Checked in, but different date")
                print(name, " Found in OK file, but different date")
                rowm += 1
                break
            else:
                good += 1
                print(name," everything match!")
                break

    if found == False:
        # print(name, " Can't find customer Name")
        not_found += 1
        worksheet_nf.write(row_nf, 0, name)
        row_nf += 1
        print(name, " customer not found")

print('finished')
print("good:", good)
print("match: ", match)
print("canceled: ", canceled)
print("not found: ", not_found)
workbook.close()
workbook_nf.close()

