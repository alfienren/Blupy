Calculations and Definitions
============================

Below is a list of the metrics used in reporting and how they are defined. How these metrics are
calculated in the routine will also be displayed.

General Definitions
-------------------

+----------------+-------------------------------------------------------------+
|     Metric     |                          Definition                         |
+================+=============================================================+
| Postpaid Plans | Count of Plans associated with a Device                     |
|                | A Postpaid Plan must have a device and                      |
|                | a plan                                                      |
+----------------+-------------------------------------------------------------+
| Prepaid Plans  | Count of Devices without an associated Plan or String       |
+----------------+-------------------------------------------------------------+
| Services       | Count of service strings in Service (string) column in CFV  |
|                | report                                                      |
+----------------+-------------------------------------------------------------+
| Devices        | Count of device strings in Device (string) column in CFV    |
|                | report                                                      |
+----------------+-------------------------------------------------------------+
| Plans          | Count of plan strings in Plan (string) column in CFV report |
+----------------+-------------------------------------------------------------+
| Accessories    | Count of accessory strings in Accessory (string) column in  |
|                | CFV report                                                  |
+----------------+-------------------------------------------------------------+
| Add-a-Line     | Count of serive strings in Service (string) column          |
|                | containing 'ADD'                                            |
+----------------+-------------------------------------------------------------+
| Activations    | Plans + Add-a-Lines                                         |
+----------------+-------------------------------------------------------------+
| Orders         | One order per OrderNumber (string) in CFV report            |
+----------------+-------------------------------------------------------------+

Campaign Specific Definitions
-----------------------------

+-----------------+-----------------------------+------------------------------------------------------------------+
|     Campaign    |            Metric           |                            Definition                            |
+=================+=============================+==================================================================+
| DDR, Demand Gen | Estimated Gross Adds (eGAs) | Count of Devices with 50% view-through credit                    |
+-----------------+-----------------------------+------------------------------------------------------------------+
| DDR             | Prepaid Gross Adds          | Count of Devices with prepaid subcategory                        |
+-----------------+-----------------------------+------------------------------------------------------------------+
| DDR             | Postpaid Gross Adds         | Count of Devices with postpaid subcategory                       |
+-----------------+-----------------------------+------------------------------------------------------------------+
| DDR             | Prepaid SIMs                | Set of Prepaid Gross Adds with product category of SIM only      |
+-----------------+-----------------------------+------------------------------------------------------------------+
| DDR             | Postpaid SIMs               | Set of Postpaid Gross Adds with product category of SIM only     |
+-----------------+-----------------------------+------------------------------------------------------------------+
| DDR             | Prepaid Mobile Internet     | Set of Prepaid Gross Adds with Mobile Internet product category  |
+-----------------+-----------------------------+------------------------------------------------------------------+
| DDR             | Postpaid Mobile Internet    | Set of Postpaid Gross Adds with Mobile Internet product category |
+-----------------+-----------------------------+------------------------------------------------------------------+
| DDR             | Prepaid Phone               | Set of Prepaid Gross Adds with Smartphone product category       |
+-----------------+-----------------------------+------------------------------------------------------------------+
| DDR             | Postpaid Phone              | Set of Postpaid Gross Adds with Smartphone product category      |
+-----------------+-----------------------------+------------------------------------------------------------------+
| DDR             | New Device                  | Count of Devices with New TMO Order Confirmation floodlight      |
+-----------------+-----------------------------+------------------------------------------------------------------+
| DDR             | Total Gross Adds            | Prepaid Gross Adds + Postpaid Gross Adds                         |
+-----------------+-----------------------------+------------------------------------------------------------------+
| DDR             | Prepaid Gross Adds          | Prepaid SIMs + Prepaid Mobile Internet + Prepaid Phones          |
+-----------------+-----------------------------+------------------------------------------------------------------+
| DDR             | Postpaid Gross Adds         | Postpaid SIMs + Postpaid Mobile Internet + Postpaid Phones       |
+-----------------+-----------------------------+------------------------------------------------------------------+

Calculation Methods and Examples
--------------------------------

**Plans**, **Devices**, **Services** and **Accessories** are tallied similiarly::

    cfv['Plans'] = cfv['Plan (string)'].str.count(',') + 1
    cfv['Devices'] = cfv['Device (string)'].str.count(',') + 1
    cfv['Services'] = cfv['Service (string)'].str.count(',') + 1
    cfv['Accessories'] = cfv['Accessory (string)'].str.count(',') + 1

Each Plan, Device, Service, etc., is separated by a comma, therefore, the metrics are calculated by counting the number
of commas in the cell and adding 1. If a cell is blank, it is assigned as NaN (not a number) and is not counted.

**Examples:**

* Plan (string): SCWTT5MI, SCWTT5MI = 2 Plans
* Device (string):  MG552LL/A,MG542LL/A,610214633651,610214694980 = 4 Devices
* Accessory (string): blank = 0 Accessories

**Add-a-Lines** are calculated by counting the number of occurrences of the string 'ADD' in Service (string) column::

	cfv['Add-a-Line'] = cfv['Service (string)'].str.count('ADD')

**Examples:**

* Service (string) 1 = ADDUTTVNC,V500DATA,V500DATA,V500DATA = 1 Add-a-Line
* Service (string) 2 = SCUNL4HHS,ADDUTTVNC,V2HDATA,V2HDATA,ADDUTTVNC = 2 Add-a-Lines
* Service (string) 3 = Visual Voicemail,2.5 GB High-Speed Data = 0 Add-a-Lines

**Postpaid Plans** are defined as a Device with an associated Plan::

	cfv['Postpaid Plans'] = np.where(cfv['Plans'] == cfv['Devices'], cfv['Plans'],
                                     pd.concat([cfv['Plans'], cfv['Devices']], axis=1).min(axis=1))

The calculation states if the number of Plans is equal to the number of Devices, Plans are therefore equal
to Postpaid Plans. If the count of Plans and Devices is not equal, the lower count would then be equal to 
the amount of Postpaid Plans.

**Examples:**

+-------------------------------------------+-----------------------------------------+----------------+
|               Plan (string)               |             Device (string)             | Postpaid Plans |
+===========================================+=========================================+================+
| SIMPLE CHOICE PLAN: UNLIMITED TALK + TEXT | Apple iPhone 6 Plus - Gold - 64GB       |              1 |
+-------------------------------------------+-----------------------------------------+----------------+
| SCMIMDG5,SCMIMDG5,SCMIMDG5,SCMIMDG5       | MH2W2LL/A,MH2W2LL/A,MH2W2LL/A,MH2W2LL/A |              4 |
+-------------------------------------------+-----------------------------------------+----------------+
| SCWTT5MI,SCWTT5MI                         | 61021463899161                          |              1 |
+-------------------------------------------+-----------------------------------------+----------------+

**Prepaid Plans** are defined as the number of Devices without an associated Service and Plan::

    cfv['Prepaid Plans'] = np.where((cfv['Plans'] == 0) & (cfv['Services'] == 0), cfv['Devices'],
                                    np.where(cfv['Devices'] > (cfv['Plans'] & cfv['Services']),
                                             cfv['Devices'] - pd.concat([cfv['Plans'], cfv['Services']], axis=1).max(axis=1), 0))

The logic first checks if the Plan and Service cells are both 0. If so, Prepaid Plans would then be the count of
Devices. If Plans and Services are not 0, Prepaid Plans are determined by subtracting the count of Plans or Services, whichever
is largest.

**Examples:**

+---------------+-----------------------------------------------+------------------------+---------------+
| Plan (string) |                Service (string)               |    Device (string)     | Prepaid Plans |
+===============+===============================================+========================+===============+
| FUTTVN        | V2HDATA,V2HDATA,V2HDATA,SCUNL4HHS             | 610214633644,MG562LL/A |             0 |
+---------------+-----------------------------------------------+------------------------+---------------+
| SCFUTT4       | ADDUTTVNC                                     | MG562LL/A,MG562LL/A    |             1 |
+---------------+-----------------------------------------------+------------------------+---------------+
| SCFUTTBD      | MG562LL/A,MG552LL/A,610214633644,610214633651 | None                   |             3 |
+---------------+-----------------------------------------------+------------------------+---------------+

**Activations** are just the sum of Add-a-Lines and Plans::

	cfv['Activations'] = cfv['Plans'] + cfv['Add-a-Line']




