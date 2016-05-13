from xlwings import Range, Sheet
import pandas as pd
import numpy as np
from reporting import categorization, datafunc
from outputs import pacing


def flat_rate_corrections():
    planned = pacing.open_planned_media_report()
    planned = categorization.sites(planned)

    planned = planned[planned['Placement Cost Structure'].str.contains('Flat Rate') == True]

    if Range('data', 'B1').value == 'Date':
        dates = pd.DataFrame(Range('data', 'B1').vertical.value, columns=['Date'])
        dates.drop(0, inplace=True)
    else:
        dates = pd.DataFrame(Range('data', 'E1').vertical.value, columns=['Date'])
        dates.drop(0, inplace=True)

    max_date = dates['Date'].max()

    planned['Ended'] = np.where(pd.to_datetime(planned['Placement End Date']) <= max_date, 'Ended', '')

    Sheet('Flat_Rate').clear_contents()
    datafunc.chunk_df(planned, 'Flat_Rate', 'A1')
