import pandas as pd
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


print("begin comparison")



#############################################################
# check conditions and print out
def same_name(booking_name, hotel_name):
    if hotel_name[:5] == 'name\n':
        hotel_name = hotel_name[5:]

    hotel_name = hotel_name.split(",")
    hotel_name = hotel_name[1][1:] + " " + hotel_name[0]

    if hotel_name == booking_name:
        print("h:", hotel_name, "b:", booking_name)
        return True
    else:
        return False


def same_price(booking_price, hotel_price, arrival, departure):
    print("input: ", booking_price, hotel_price, arrival, departure)

    if departure[:2] == "De":
        departure = departure[9:]
        arrival = arrival[7:]
        hotel_price = hotel_price[4:]

    date_format = "%m/%d/%Y"
    a = datetime.strptime(departure, date_format)
    b = datetime.strptime(arrival, date_format)
    delta = a - b
    print("hotel price: ", float(delta.days) * float(hotel_price), "booking price: ", float(booking_price))
    print(float(delta.days) * float(hotel_price) == float(booking_price))
    if (float(delta.days) * float(hotel_price) == float(booking_price)):
        return True


# check loop
for index, row in b_df.iterrows():

    # initialzie what to compare
    name = row['Guest name(s)']
    check_in_date = row['Check-in']
    check_out_date = row['Check-out']
    price = row['Price'][:-3]
    found = False

    # check if in ok, but different price
    for index_ok, row_ok in o_df.iterrows():
        if (same_name(name.lower(), row_ok['Name'].lower())):
            print(name + "**MATCH***********************************************")
            found = True
            if same_price(price, row_ok['Rate'][1:], row_ok['Arrival'], row_ok['Departure']):
                print(price + "**MATCH***********************************************")

            else:
                print(name, price, "found in Ok file ,but different price")

    if (found == False):
        # check if in canceled
        for index_c, row_c in c_df.iterrows():
            if (same_name(name.lower(), row_c['Name'].lower())):
                found = True
                print(name + "**BAD***********************************************")
