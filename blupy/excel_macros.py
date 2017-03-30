from analytics.data import advertisers, api, streams
from analytics.reporting import dashboards, qa, ias
from dcm import advertiser, floodlights, pixels


def tmobile_weekly_reporting():
    advertisers.UpdateAdvertisers().tmo()


def metro_weekly_reporting():
    advertisers.UpdateAdvertisers().metro()


def wfm_weekly_reporting():
    advertisers.UpdateAdvertisers().wfm()


def cross_channel_dashboard():
    dashboards.CrossChannel().generate_dashboard()


def tmo_output_flat_rates():
    qa.QA().flat_rates()


def site_pacing_report():
    qa.QA().site_pacing()

############################## Trafficking ####################################

def create_new_floodlights():
    floodlights.Floodlights().insert()


def get_floodlight_list():
    floodlights.Floodlights().get()


def all_floodlight_tags():
    floodlights.Floodlights().generate_all_tags()


def listed_floodlight_tags():
    floodlights.Floodlights().generate_selected_tags()


def get_pixel_list():
    pixels.Pixels().get()


def piggyback_pixels():
    pixels.Pixels().implement()


def delete_pixels():
    pixels.Pixels().delete()


def list_campaigns():
    advertiser.Advertiser().list_campaign_names_ids()

############################## IAS ####################################

def merge_ias():
    ias.IASReporting().download_data()

############################## Placed ####################################

def placed_data():
    api.ReportingAPI().placed()


def cfv_report():
    streams.DatoramaStreams().custom_floodlights()