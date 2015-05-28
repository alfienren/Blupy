import pandas as pd

def messaging(data):

    # Message Bucket, Category and Offer
    # 90% of the time, the message bucket, category and offer can be determined from the creative field 1 column. It
    # follows a pattern of Bucket_Category_Offer
    data['Creative Field 1'] = data['Creative Field 1'].str.replace('Creative Type: ', '')
    data['Creative Field 1'] = data['Creative Field 1'].str.replace('(', '')
    data['Creative Field 1'] = data['Creative Field 1'].str.replace(')', '')

    # If Creative Field 1 is equal to (not set), this is either a 1x1 or a placement with logo creative. (not set)
    # fields are therefore set as 'TMO Unique Creative', which is how this has been handled historically.
    data['Creative Field 1'] = data['Creative Field 1'].str.replace('not set', 'TMO Unique Creative')

    # Message Bucket is determined by splitting Creative Field 1 and taking the first word.
    data['Message Bucket'] = data['Creative Field 1'].str.split('_').str.get(0)

    # Message Category is determined by splitting Creative Field 1 and taking the second word.
    data['Message Category'] = data['Creative Field 1'].str.split('_').str.get(1)

    # Message Offer is determined by splitting Creative Field 1 and taking the third word. If the offer is not set,
    # it can sometimes be found in the Creative Groups 2 column. For blanks in the Message Offer column, it will try
    # to pull in the offer from the Creative Groups 2 column.
    data['Message Offer'] = data['Creative Field 1'].str.split('_').str.get(2)
    data['Message Offer'].fillna(data['Creative Groups 2'], inplace=True)

    return data