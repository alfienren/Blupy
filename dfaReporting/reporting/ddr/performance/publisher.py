import numpy as np
import pandas as pd
from xlwings import Workbook, Range

import reporting.ddr.performance.common
from reporting import paths
from reporting.ddr.performance.common import quarter_start_year, week_of, row


def publishers(dr):
    path = paths.path_select()
    week = week_of(dr)

    # Publisher Performance
    pub_dr = dr[(dr['Campaign'] == 'DR') & (dr['Date'] >= quarter_start_year())]
    pub_dr = pub_dr.groupby(['Site', 'Placement Messaging Type', 'Week', 'Date'])
    pub_dr = pd.DataFrame(pub_dr.sum()).reset_index()

    pub_dr['Tactic'] = np.where(pub_dr['Placement Messaging Type'].str.contains('FBX ') == True,
                                'FBX Remessaging', pub_dr['Placement Messaging Type'])

    # Quarter
    q_dr = pub_dr[pub_dr['Date'] >= quarter_start_year()]
    q_dr = q_dr.groupby(['Site', 'Tactic'])
    q_dr = pd.DataFrame(q_dr.sum()).reset_index()

    # Last Week
    last_week = pub_dr[pub_dr['Week'] == pub_dr['Week'].max()]
    last_week = last_week.groupby(['Site', 'Tactic'])
    last_week = pd.DataFrame(last_week.sum()).reset_index()

    last_week.rename(columns={'NET Media Cost': week}, inplace= True)

    # Publishers Overall
    sites = q_dr.groupby('Site')
    sites = pd.DataFrame(sites.sum()).reset_index()

    sites['CPO'] = sites['NET Media Cost'] / sites['Orders']

    # Brand Remessaging

    br = dr[dr['Campaign'] == 'Brand Remessaging']
    br_quarter = br[br['Date'] >= quarter_start_year(start='late')]

    br_quarter = br_quarter.groupby('Site')
    br_quarter = pd.DataFrame(br_quarter.sum().reset_index())
    br_quarter['Traffic Yield'] = br_quarter['Traffic Actions'].astype(float) / \
                                  br_quarter['Impressions'].astype(float)

    pacing_wb = Workbook(path)
    pacing_wb.set_current()

    Range(reporting.ddr.performance.common.performance_sheet(), 'A' + str(row()), index=False).value = q_dr[
        ['Site', 'Tactic', 'Orders']]
    Range(reporting.ddr.performance.common.performance_sheet(), 'E' + str(row()), index= False).value = \
        q_dr['NET Media Cost']

    Range(reporting.ddr.performance.common.performance_sheet(), 'D' + str(row()), index= False).value = last_week[week]
    Range(reporting.ddr.performance.common.performance_sheet(), 'D' + str(row() - 1)).value = week

    Range(reporting.ddr.performance.common.performance_sheet(), 'A' + str(row() + 1), index= False).value = \
        sites[['Site', 'Orders', 'CPO']]

    Range(reporting.ddr.performance.common.performance_sheet(), 'C7').value = \
        pub_dr['Week'].max().strftime('%m/%d/%Y').lstrip('0').replace(' 0', ' ')

    Range(reporting.ddr.performance.common.performance_sheet(), 'B' + str(row()), index= False).value = \
        br_quarter[['Traffic Yield', 'Impressions']]

    Range(reporting.ddr.performance.common.performance_sheet(), 'E' + str(row()), index= False).value = \
        br_quarter['Traffic Actions']

    Range(reporting.ddr.performance.common.performance_sheet(), 'F' + str(row()), index= False).value = \
        br_quarter['NET Media Cost']

    pacing_wb.save()
    pacing_wb.close()
