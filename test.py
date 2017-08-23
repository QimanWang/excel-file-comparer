import pandas as pd
import xlsxwriter
from datetime import datetime

# load and prepare the booking df
xl = pd.ExcelFile("Booking.com 7.1-8.13.xls")
booking_df = xl.parse("Sheet1", header=0, keep_default_na=False)
b_df = booking_df[booking_df['Status'] == 'ok']
for index, row in b_df.iterrows():
    if row['Guest name(s)'] == '':
        name = str(row['Booked by'])
        guest_name = name.split(",")
        b_df.loc[index, 'Guest name(s)'] = guest_name[1][1:] + " " + guest_name[0]

##########################################
# loading the ok DF
xl = pd.ExcelFile("30031-GuestListDetailed ok 7.1-8.13.xlsx")
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
xl = pd.ExcelFile("30031-GuestListDetailed cancel 7.1-8.13.xlsx")
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


#########################################
print("begin comparison")
workbook = xlsxwriter.Workbook('Expenses01.xlsx')
worksheet = workbook.add_worksheet()
rowm = 0
coln = 0


# # print b_df
# print(list(b_df))
# for index, row in b_df.iterrows():
#     print("index:::", index, row['Guest name(s)'].lower(),row['Check-in'],row['Check-out'],row['Price'])
#

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


def same_price(booking_price, hotel_price, arrival, departure):
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
    if (float(delta.days) * float(hotel_price) == float(booking_price)):
        return True


# check loop
for index, row in b_df.iterrows():

    # initialzie what to compare
    name = row['Guest name(s)']
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

            if (h_arrival == b_checkin) & (h_departure == b_checkout):
                found = True

                if not same_price(price, row_ok['Rate'][1:], row_ok['Arrival'], row_ok['Departure']):
                    print(name, price, "found in Ok file ,but different price")
                    break

    if (found == False):
        # check if in canceled
        for index_c, row_c in c_df.iterrows():
            if (same_name(name.lower(), row_c['Name'].lower())):
                print(name + "**BAD***********************************************")
                break