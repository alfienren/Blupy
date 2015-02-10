Calculations and Definitions
============================

Below is a list of the metrics used in reporting and how they are defined. How these metrics are
calculated in the routine will also be displayed.


+----------------+-------------------------------------------+---------------------------------------------+
|     Metric     |                 Definition                |                   Example                   |
+================+===========================================+=============================================+
| Postpaid Plans | Count of Plans associated with a Device   | Device1 -- Plan1 = 1 Postpaid Plan          |
|                | A Postpaid Plan must have a device        | Device2 -- Plan2 = 1 Postpaid Plan          |
|                | and a plan                                | Device3 --       = 0 Postpaid Plans         |
|                |                                           | Total: 2 Postpaid Plans                     |
+----------------+-------------------------------------------+---------------------------------------------+
| Prepaid Plans  | Count of Devices without a Plan or String |                                             |
+----------------+-------------------------------------------+---------------------------------------------+
| Services       | Count of service strings in Device        | ServiceString1, ServiceString2 = 2 Services |
|                | (string) column                           | ServiceString3                 = 1 Service  |
|                |                                           | Total: 3 Services                           |
|                |                                           |                                             |
|                |                                           |                                             |
+----------------+-------------------------------------------+---------------------------------------------+