import requests
from scripts.utils import get_api_key


API_KEY = get_api_key()
SPORT = 'basketball_ncaab'
REGIONS = 'us'
MARKETS = 'spreads,h2h'
ODDS_FORMAT = 'american'
DATE_FORMAT = 'iso'


def get_odds(from_date, to_date):
    odds_response = requests.get(
        f'https://api.the-odds-api.com/v4/sports/{SPORT}/odds',
        params={
            'api_key': API_KEY,
            'regions': REGIONS,
            'markets': MARKETS,
            'oddsFormat': ODDS_FORMAT,
            'dateFormat': DATE_FORMAT,
        }
    )
    if odds_response.status_code != 200:
        print(f"Error status code: {odds_response.status_code}")
    else:
        print('Remaining requests',
              odds_response.headers['x-requests-remaining'])
        print('Used requests', odds_response.headers['x-requests-used'])
        return odds_response.json()


def get_historical_odds(date):
    odds_response = requests.get(
        f'https://api.the-odds-api.com/v4/sports/{SPORT}/odds-history',
        params={
            'api_key': API_KEY,
            'regions': REGIONS,
            'markets': MARKETS,
            'oddsFormat': ODDS_FORMAT,
            'dateFormat': DATE_FORMAT,
            'date': date
        }
    )
    if odds_response.status_code != 200:
        print(f"Error status code: {odds_response.status_code}")
    else:
        print('Remaining requests',
              odds_response.headers['x-requests-remaining'])
        print('Used requests', odds_response.headers['x-requests-used'])
        return odds_response.json()
