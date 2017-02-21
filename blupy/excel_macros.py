from analytics.data_refresh import advertisers
from analytics.reporting import dashboards, qa
from dcm.trafficking import advertiser, floodlights, pixels
from http.urls import URLs


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


def list_campaigns():
    advertiser.Advertiser().list_campaign_names_ids()


#################################### URLs ##############################

def get_url_descriptions():
    URLs().url_descriptions()

if __name__ == 'floodlight_urls':
    URLs().list_floodlights_from_urls()

if __name__ == 'get_url_descriptions':
    URLs().url_descriptions()