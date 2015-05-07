import numpy as np

def output(data):

    # The DFA field DBM Cost is more accurate for placements using dynamic bidding. If a placement is not using
    # dynamic bidding, DBM Cost = 0. Therefore, if DBM cost does not equal 0, replace the row's media cost with
    # DBM cost. If DBM Cost = 0, Media Cost stays the same.
    data['Media Cost'] = np.where(data['DBM Cost USD'] != 0, data['DBM Cost USD'], data['Media Cost'])

    # Adjust spend to Net to Client
    data['NTC Media Cost'] = 0

    # DBM Cost column is then removed as it is no longer needed.
    data.drop('DBM Cost USD', 1, inplace=True)

    # Create week column by taking the oldest date in the data
    data['Week'] = data['Date'].min()

    # Add columns for Video Completions and Views, primarily for compatibility between campaigns that run video and
    # don't run video. Those that don't can keep the columns set to zero, but those that have video can then be
    # adjusted with passback data.
    data['Video Completions'] = 0
    data['Video Views'] = 0

    return data