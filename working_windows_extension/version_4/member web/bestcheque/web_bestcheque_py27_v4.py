import pandas as pd
import xlsxwriter
import os
import datetime

# find the files needed to process
hotel_file = ""
website_file = ""
files = [f for f in os.listdir('.') if os.path.isfile(f)]
for f in files:
    if "member" in f.lower() and f[:2] != "~$" and "-" in f:
        hotel_file = f
    if 'bestcheque' in f.lower() and f[:2] != "~$" and f[-4:] == 'xlsx':
        website_file = f

msg = "Files detected:\n" + website_file + "\n" + hotel_file + '\n' + "Press Enter key to begin the process."
a = raw_input(msg)


# define functions
def same_name(booking_name, hotel_name):
    if hotel_name[:5] == 'name\n':
        hotel_name = hotel_name[5:]

    if ',' in hotel_name:
        hotel_name = hotel_name.split(",")
        hotel_name = hotel_name[1][:] + " " + hotel_name[0]

    # if '/' in booking_name:
    #     booking_name = booking_name.split("/")
    #     booking_name = booking_name[1][:] + " " + booking_name[0]

    # print("h:", hotel_name, "b:", booking_name)
    if hotel_name == booking_name:

        return True
    else:
        return False


######################################################################################
# load the web file
xl = pd.ExcelFile(website_file)
web_df = xl.parse(header=0, keep_default_na=False)

##########################################
# loading the hotel df
xl = pd.ExcelFile(hotel_file)
hotel_df = xl.parse(header=0, keep_default_na=False, )
print(list(hotel_df))

########################################################################
#
match = 0
canceled = 0
not_found = 0

print("begin comparison")

# create worksheet
workbook = xlsxwriter.Workbook('Expenses.xlsx')
worksheet = workbook.add_worksheet()
rowm = 0
coln = 0
# worksheet.write(rowm, coln, "Confirmation Number")
worksheet.write(rowm, coln + 0, "Conf Number")
worksheet.write(rowm, coln + 1, "Guest Name")
worksheet.write(rowm, coln + 2, "Price")
worksheet.write(rowm, coln + 3, "Description")
rowm += 1

workbook_nf = xlsxwriter.Workbook('Customers Not found.xlsx')
worksheet_nf = workbook_nf.add_worksheet()
row_nf = 0

###################################################################
# check loop
good = 0
status_col_name_index = list(hotel_df).index("GTD")
status_col_name = list(hotel_df)[status_col_name_index + 1]
print("status", status_col_name)
for index, row in web_df.iterrows():

    # initialzie what to compare
    name = row['Guest Name']
    check_in_date = row['Arrival Date']
    check_out_date = row['Depart Date']
    # print('check_in_date = ', row['Arrival Date'])
    # print('check_out_date =', row['Depart Date'])

    b_checkin = check_in_date.to_pydatetime()
    b_checkout = check_out_date.to_pydatetime()

    # print(a)
    found = False

    # check if in ok, but different price
    for index_ok, row_ok in hotel_df.iterrows():
        # bestcheque: not good, hotel:not good
        # print('bestcheque',name)
        # print('hotel',row_ok['Guest Name'].lower())
        if '/' in name:
            name = name.split("/")
            name = name[1][:] + " " + name[0]

        if (same_name(name.lower(), row_ok['Guest Name'].lower())):
            found = True

            if row_ok[status_col_name] != "":
                print(name, "ROOM canceled", row_ok[status_col_name])

                # worksheet.write(rowm, coln, "")
                worksheet.write(rowm, coln + 0, row_ok['Conf/Cxl#'])
                worksheet.write(rowm, coln + 1, name)
                worksheet.write(rowm, coln + 2, row['Room Rev'])
                worksheet.write(rowm, coln + 3, "Cancelled")

                print(name, " found in CANCEL file")
                rowm += 1
                canceled += 1
                break

            # if not found in cancel file, we check the date
            # hotel date

            date_format = "%m/%d/%Y"
            a = str(row_ok['Arrival']).split(" ")
            b = a[0].split("-")
            arrival = "/".join([b[1], b[2], b[0]])
            # print(arrival)

            h_arrival = datetime.datetime.strptime(arrival, date_format)
            h_departure = h_arrival + datetime.timedelta(days=row_ok['Nts'])

            # booking.com date
            # b_checkin = datetime.datetime(int(check_in_date[0]), int(check_in_date[1]), int(check_in_date[2][:2]))
            # b_checkout = datetime.datetime(int(check_out_date[0]), int(check_out_date[1]),
            #                                int(check_out_date[2][:2]))
            found = True
            match += 1

            # print('h_arrival: ', h_arrival)
            # print('h_departure: ', h_departure)
            # print("b_checkin: ", b_checkin)
            # print("b_checkout", b_checkout)

            if not ((h_arrival == b_checkin) & (h_departure == b_checkout)):
                # worksheet.write(rowm, coln, "")
                worksheet.write(rowm, coln + 0, row_ok['Conf/Cxl#'])
                worksheet.write(rowm, coln + 1, name)
                worksheet.write(rowm, coln + 2, row['Room Rev'])
                worksheet.write(rowm, coln + 3, "Date changed")
                print(name, " Found in OK file, but different date")
                rowm += 1
                break
            else:
                good += 1
                print(name, " everything match!")
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
