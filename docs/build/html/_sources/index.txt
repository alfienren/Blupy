Weekly Reporting Documentation
========

Introduction
------------

Campaign reports from DFA are pulled weekly to provide the media and analytics teams 
insight into performance of running campaigns. Campaign reporting consists of two 
reports from DFA, Site Activity and Custom Floodlight Variables. These reports are
combined and then munged in order to provide cleaned, readable data on performance.
To speed up the process of pulling reports and munging reports, a routine was created
using VBA and Python. This documentation will detail the code and the logic behind how
calculations are made.

Table of Contents
-----------------

  .. toctree::
  	 :maxdepth: 2

  	 weeklyreporting
  	 calculations
  	 indices

Source
------

I keep the working code on Github to better maintain changes and track issues. No
private data is included. The project can be found here: https://github.com/TMOptimedia/Weekly_DFA_Reporting.
If you would like to contribute by adding issues or submitting pull requests, please email me 
and I will add you to the organization.