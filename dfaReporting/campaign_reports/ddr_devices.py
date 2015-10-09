import pandas as pd
from xlwings import Range

def device_feed():

    device_feed_path = Range('Action_Reference', 'AE1').value

    device_lookup = pd.read_table(device_feed_path)

    return device_lookup

def top_15_devices():

    