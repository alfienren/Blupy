import datetime

import arrow
import numpy as np
import pandas as pd
from xlwings import Range


class Categorization(object):

    def __init__(self):
        self.spanish_campaigns = '|'.join(list(['Spanish', 'Hispanic', 'SL', 'Latino', 'Univision', 'Telemundo']))
        self.quarter = 'Q2'

    @staticmethod
    def placements(data, adv='tmo'):
        desktop, mobile, video, standard, rm, custom, rem, social = [], [], [], [], [], [], [], []

        match_strings = [desktop, mobile, video, standard, rm, custom, rem, social]

        col = chr(ord('A'))

        for i in range(0, len(match_strings)):
            ref = Range('Lookup', col + '1').vertical.value
            match_strings[i].append('|'.join(list(ref)))
            col = chr(ord(str(col)) + 1)

        tmob = '|'.join(list(['T-Mobile', 'T-Mob']))

        data['Placement2'] = np.where(data['Placement'].str.contains(tmob) == True,
                                      data['Placement'].str.replace(tmob, ''), data['Placement'])

        data['Platform'] = np.where((data['Placement2'].str.contains(mobile[0L]) == True), 'Mobile',
                                    np.where(data['Placement2'].str.contains(desktop[0L]) == True, 'Desktop',
                                             'Desktop'))

        data['Creative2'] = np.where(data['Placement2'].str.contains(video[0L]) == True, 'Video',
                                     np.where(data['Placement'].str.contains(standard[0L]) == True, 'Display',
                                              'Display'))

        data['Creative3'] = np.where(data['Placement2'].str.contains(rm[0L]) == True, 'Rich Media',
                                     np.where(data['Placement2'].str.contains(custom[0L]) == True, 'Custom',
                                              np.where(data['Placement2'].str.contains(rem[0L]) == True, 'Remessaging',
                                                       np.where(data['Placement2'].str.contains(social[0L]) == True,
                                                                'Social', 'Standard'))))

        if adv == 'tmo':
            data['Category'] = data['Platform'] + ' - ' + data['Creative2'] + ' - ' + data['Creative3']
            data['Category_Adjusted'] = data['Platform'] + ' - ' + data['Creative2']

        elif adv == 'metro':
            data['TMO_Category'] = data['Platform'] + ' - ' + data['Creative2'] + ' - ' + data['Creative3']
            data['TMO_Category_Adjusted'] = data['Platform'] + ' - ' + data['Creative2']

            cat = pd.DataFrame(Range('Lookup', 'J1').table.value, columns=Range('Lookup', 'J1').horizontal.value)
            cat.drop(0, inplace=True)
            cat.drop_duplicates(subset='Placement Name', inplace=True)

            data = pd.merge(data, cat, left_on='Placement', right_on='Placement Name', how='left')
            data.drop('Placement Name', axis=1, inplace=True)

            data.drop_duplicates(inplace=True)

        return data

    @staticmethod
    def creative(data):
        # Message Bucket, Category and Offer
        # 90% of the time, the message bucket, category and offer can be determined from the creative field 1 column. It
        # follows a pattern of Bucket_Category_Offer
        data['Creative Field 1'] = data['Creative Field 1'].str.replace('Creative Type: ', '')
        data['Creative Field 1'] = data['Creative Field 1'].str.replace('(', '')
        data['Creative Field 1'] = data['Creative Field 1'].str.replace(')', '')

        # If Creative Field 1 is equal to (not set), this is either a 1x1 or a placement with logo creative. (not set)
        # fields are therefore set as 'TMO Unique Creative', which is how this has been handled historically.

        # Message Bucket is determined by splitting Creative Field 1 and taking the first word.
        data['Message Bucket'] = data['Creative Field 1'].str.split('_').str.get(0)

        # Message Category is determined by splitting Creative Field 1 and taking the second word.
        data['Message Category'] = data['Creative Field 1'].str.split('_').str.get(1)

        # Message Offer is determined by splitting Creative Field 1 and taking the third word. If the offer is not set,
        # it can sometimes be found in the Creative Groups 2 column. For blanks in the Message Offer column, it will try
        # to pull in the offer from the Creative Groups 2 column.

        data['Creative Bucket'] = data['Creative Field 1'].str.split('_').str.get(2)
        data['Creative Bucket'].fillna(data['Creative Field 1'], inplace=True)

        data['Creative Theme'] = data['Creative Groups 2']

        return data

    @staticmethod
    def sites(data):
        site_ref = pd.DataFrame(Range('Lookup', 'Q1').table.value, columns = Range('Lookup', 'Q1').horizontal.value)
        site_ref.drop(0, inplace=True)
        site_ref.drop_duplicates(subset='DFA', inplace=True)

        data = pd.merge(data, site_ref, left_on= 'Site (DCM)', right_on= 'DFA', how= 'left')
        data.drop('DFA', axis = 1, inplace = True)

        return data

    def media_plan(self, data):
        data['Campaign2'] = np.where(data['Campaign'].str.contains('BidManager') == True, data['Placement'],
                                     data['Campaign'])

        data['Media Plan'] = np.where(data['Campaign2'].str.contains('Brand Remessaging|Brand Rms') == True,
                                      'Brand Remessaging',
                                      np.where(data['Campaign2'].str.contains('DDR') == True, 'DDR',
                                               np.where(data['Campaign2'].str.contains('Forbes') == True,
                                                        'Forbes Sponsorship',
                                                        np.where(data['Campaign2'].str.contains('FEP|ADU|OXYGEN') == True,
                                                                 'FEP Upfront',
                                                                 np.where(data['Campaign2'].str.contains(
                                                                     'DemandGen|Demand Gen') == True, 'Demand Gen',
                                                                          np.where(data['Campaign2'].str.contains(
                                                                              'Network') == True, 'Network',
                                                                                   'Super Bowl'))))))

        data['Media Plan'] = np.where(data['Campaign'].str.contains(self.spanish_campaigns) == True,
                                      'SL ' + data['Media Plan'],
                                      data['Media Plan'])

        data.drop('Campaign2', axis=1, inplace=True)

        return data


    @staticmethod
    def message_campaign(data):
        message_table = pd.DataFrame(Range('Lookup', 'T1').table.value, columns=Range('Lookup', 'T1').horizontal.value)
        message_table.drop(0, inplace=True)

        message_table.drop_duplicates(inplace=True)

        data_merged = pd.merge(data, message_table, on=['Media Plan', 'Creative Groups 2'], how='left')

        return data_merged


    def language(self, data):
        data['Language'] = np.where(data['Campaign'].str.contains(self.spanish_campaigns) == True, 'SL', 'EL')

        return data

    @staticmethod
    def mondays(dates):
        monday = dates - datetime.timedelta(days= dates.weekday()) + datetime.timedelta(days= 7, weeks= -1)

        return monday

    @staticmethod
    def date_columns(data):
        quarters = {
            'January': 'Q1',
            'February': 'Q1',
            'March': 'Q1',
            'April': 'Q2',
            'May': 'Q2',
            'June': 'Q2',
            'July': 'Q3',
            'August': 'Q3',
            'September': 'Q3',
            'October': 'Q4',
            'November': 'Q4',
            'December': 'Q4'
        }

        data['Date2'] = pd.to_datetime(data['Date'])

        data['Month'] = data['Date2'].apply(lambda x: arrow.get(x).format('MMMM'))
        data['Quarter'] = data['Month'].apply(lambda x: quarters[x])
        data['Week'] = data['Date2'].apply(lambda x: Categorization.mondays(x))
        data.drop('Date2', axis = 1, inplace = True)

        return data

    @staticmethod
    def dr_sites():
        site = '|'.join(list(['AOD', 'ASG', 'Amazon', 'Bazaar Voice', 'eBay', 'Magnetic', 'Yahoo']))

        return site

    @staticmethod
    def quarter_start_year(start='quarter_start'):
        if start != 'quarter_start':
            quarter = '4/18/2016'
        else:
            quarter = '4/1/2016'

        return quarter

    @staticmethod
    def dr_placement_message_type(data):
        message_type = pd.DataFrame(Range('Lookup', 'K3').table.value, columns = Range('Lookup', 'K3').horizontal.value)
        message_type.drop(0, inplace= True)
        message_type.drop_duplicates(keep='last', inplace=True)

        data = pd.merge(data, message_type, left_on= 'Placement', right_on= 'Placement_category', how= 'left')
        data.drop(['Campaign2', 'Placement_category'], axis= 1, inplace= True)

        return data

    @staticmethod
    def dr_tactic(data):
        tactic = pd.DataFrame(Range('Lookup', 'AD1').table.value, columns = Range('Lookup', 'AD1').horizontal.value)
        tactic.drop(0, inplace=True)
        tactic.drop_duplicates(keep='last', inplace=True)

        data = pd.merge(data, tactic, on='Placement Messaging Type', how='left')

        return data

    @staticmethod
    def dr_creative_categories(data):
        categories = pd.DataFrame(Range('Lookup', 'Y1').table.value,
                                           columns = Range('Lookup', 'Y1').horizontal.value)
        categories.drop(0, inplace=True)
        categories.drop_duplicates(keep='last', inplace=True)

        data = pd.merge(data, categories, on='Creative Groups 2', how='left')

        return data

    @staticmethod
    def search_bucket_class(data):
        lookup_table = pd.DataFrame(Range('Lookup', 'A1').table.value,
                                    columns=Range('Lookup', 'A1').horizontal.value)
        lookup_table.drop(0, inplace=True)
        lookup_table.drop_duplicates(keep='last', inplace=True)

        data = pd.merge(data, lookup_table, how='left', on='Bucket Class')

        return data

    @staticmethod
    def search_cfv_categories(data):
        web_team = '|'.join(list(['DR-Brand', 'Remarketing', 'PLAs', 'Affiliate', 'Whistleout']))

        data['Pre vs. Post'] = data['Site (DFA)'] + ' ' + data['Product Subcategory']
        data['Web Team Post/Pre Paid'] = np.where(data['Bucket Class'].str.match(web_team) == True,
                                                  data['Site (DFA)'] + data['Pre vs. Post'] + ' Web', None)

        data['Bucket Class Pre vs. Post'] = data['Bucket Class'] + ' ' + data['Product Subcategory']
        data['Web Team'] = np.where(data['Bucket Class'].str.match(web_team) == True, data['Site (DFA)'] + ' Web', None)

        return data


    @staticmethod
    def categorize_report(data, adv='tmo'):
        data = Categorization().sites(data)

        if adv == 'tmo':
            data = Categorization().placements(data, adv='tmo')
            data = Categorization().creative(data)
            data = Categorization().dr_placement_message_type(data)
            data = Categorization().media_plan(data)
            data = Categorization().message_campaign(data)
        else:
            data = Categorization().placements(data, adv='metro')

        data = Categorization().language(data)
        data = Categorization().date_columns(data)

        return data