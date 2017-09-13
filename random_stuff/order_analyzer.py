import sys
import pandas as pd
import re

order_data = pd.read_csv(sys.argv[1])
print(order_data.head())

# check state


def validate_state(state):
    invalid_states = ['CT', 'ID', 'IL', 'MA', 'NJ', 'OR', 'PA']
    if state in invalid_states:
        return False
    else:
        return True

# check zipcode


def validate_zipcode(zipcode):
    if len(str(zipcode)) == 5 or len(str(zipcode)) == 9:
        if sum(int(digit) for digit in str(zipcode)) <= 20:
            return True
        else:
            return False
    else:
        return False

# check email


def validate_email(email):
    email_format = re.compile('^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$')
    if email_format.match(email):
        return True
    else:
        return False


# iterate through every row to calculate validity and print
for row in range(0, len(order_data)):
    if (validate_state(order_data.loc[row, 'state']) &
        validate_zipcode(order_data.loc[row, 'zipcode']) &
            validate_email(order_data.loc[row, 'email'])):
        print(order_data.loc[row, 'name'], 'valid')
    else:
        print(order_data.loc[row, 'name'], 'invalid')
