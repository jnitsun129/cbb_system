from datetime import datetime
import pandas as pd
from scripts.teams_dict import TEAMS_DICT
from scripts.utils import get_dates, print_plays
from scripts.kenpom import fetch_kenpom, Driver
from scripts.odds_api import get_odds
from scripts.spreadsheet import populate_spreadsheet
from scripts.dynamo import upload_plays, get_plays


def calculate_system(games: list, kenpom: pd.DataFrame, date: str, book_name='fanduel') -> dict:
    system_games = {}
    plays = None
    last_game = 0
    try:
        plays = get_plays(date)['data']
    except:
        pass
    num_games = 1
    for game in games:
        home_team = game['home_team']
        away_team = game['away_team']
        odds = None
        for book in game['bookmakers']:
            if book['key'] == book_name:
                odds = book['markets']
                break
        other_book_flag = False
        other_book_name = None
        if odds is None:
            try:
                other_book = game['bookmakers'][0]
            except IndexError:
                breakpoint()
            odds = other_book['markets']
            other_book_flag = True
            try:
                other_book_name = other_book['Title']
            except:
                other_book_name = other_book['title']
        moneyline_odds = None
        if len(odds) == 2:
            for moneyline in odds[0]['outcomes']:
                if moneyline['price'] < 0:
                    moneyline_odds = moneyline['price']
            for point_spread in odds[1]['outcomes']:
                if point_spread['point'] < 0:
                    favorite = point_spread['name']
                    spread_odds = point_spread['price']
        else:
            for point_spread in odds[0]['outcomes']:
                try:
                    if point_spread['point'] < 0:
                        favorite = point_spread['name']
                        spread_odds = point_spread['price']
                except KeyError:
                    breakpoint()

        if favorite == home_team:
            to_add = True
            try:
                home_team = kenpom[kenpom['Team'] ==
                                   TEAMS_DICT[home_team]].iloc[0]
            except KeyError as e:
                print(home_team)
                break
            try:
                away_team = kenpom[kenpom['Team'] ==
                                   TEAMS_DICT[away_team]].iloc[0]
            except KeyError as e:
                print(away_team)
                break
            try:
                home_rank = int(home_team['Rk'])
            except TypeError as e:
                print(home_team)
                break
            try:
                away_rank = int(away_team['Rk'])
            except TypeError as e:
                print(away_team)
                break
            if to_add and home_rank > away_rank:
                game_dict = {"GAME_NUMBER": num_games, 'home_team': home_team.to_dict(
                ), 'away_team': away_team.to_dict(), 'spread': abs(point_spread['point']) * -1, 'book': (book_name.capitalize() if not other_book_flag else other_book_name), 'home_team_score': 0, 'away_team_score': 0, 'id': game['id'], 'ml_odds': moneyline_odds, 'spread_odds': spread_odds}
                add_flag = True
                if plays is not None and len(plays) > 0:
                    system_games = plays
                    for _, value in plays.items():
                        if value['home_team']['Team'] == game_dict['home_team']['Team']:
                            add_flag = False
                            break
                    if add_flag:
                        game_dict['GAME_NUMBER'] = last_game + num_games
                        system_games[last_game + num_games] = game_dict
                        num_games += 1
                else:
                    system_games[num_games] = game_dict
                    num_games += 1
    return system_games


def run(date: datetime) -> None:
    date_str = date.strftime("%Y-%m-%d")
    driver = Driver()
    kenpom = fetch_kenpom(driver)
    print("Kenpom Data Fetched")
    driver.quit()
    dates = get_dates(date)
    games = get_odds(dates['from'], dates['to'])
    print("API Call Made")
    plays = calculate_system(games, kenpom, date_str)
    upload_plays(plays, date_str)
    print(f"{date_str} plays uploaded to dynamo")
    populate_spreadsheet(date_str, plays)
    print("Google Sheet Populated")
    print_plays(plays, date_str)
