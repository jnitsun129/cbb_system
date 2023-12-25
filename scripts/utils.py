from datetime import datetime
import csv


def todays_date() -> datetime.date:
    return datetime.now()


def num_pages() -> int:
    target_date = datetime(2023, 12, 6)
    current_date = datetime.now()
    delta = current_date - target_date
    return delta.days


def get_dates(date: datetime) -> dict:
    beginning_of_day = date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_of_day = end_of_day = beginning_of_day.replace(
        hour=23, minute=59, second=59, microsecond=999999)
    formatted_beginning_of_day = beginning_of_day.strftime(
        '%Y-%m-%dT%H:%M:%SZ')
    formatted_end_of_day = end_of_day.strftime('%Y-%m-%dT%H:%M:%SZ')

    return {'from': formatted_beginning_of_day, 'to': formatted_end_of_day}


def get_spreadsheet_id() -> str:
    with open('./utility_files/sheet_id.txt', 'r') as file:
        for line in file:
            return line.strip()


def format_date(input_date_str: str) -> str:
    try:
        input_date = datetime.strptime(input_date_str, "%Y-%m-%d")
        formatted_date = input_date.strftime("%B %d, %Y")
        day = input_date.day
        if day in [1, 21, 31]:
            formatted_date = formatted_date.replace(str(day), f"{day}st")
        elif day in [2, 22]:
            formatted_date = formatted_date.replace(str(day), f"{day}nd")
        elif day in [3, 23]:
            formatted_date = formatted_date.replace(str(day), f"{day}rd")
        else:
            formatted_date = formatted_date.replace(str(day), f"{day}th")
        return formatted_date
    except ValueError:
        return "Invalid date format"


def print_plays(plays: list, date: str) -> None:
    print(f'{len(plays)} System Plays for {format_date(
        date)} (All odds are fanduel)\n\n\n')


def get_aws_credentials() -> dict:
    d = {}
    with open('./utility_files/rootkey.csv') as file:
        csv_reader = csv.reader(file, delimiter=',')
        for index, row in enumerate(csv_reader):
            if index != 0:
                d['access_key'] = row[0]
                d['secret_key'] = row[1]
    d['region_name'] = 'us-east-2'
    return d


def get_api_key() -> str:
    with open("./utility_files/odds_api_key.txt", 'r') as file:
        for line in file:
            line = line.strip()
            return line


def get_chromedriver_path() -> str:
    with open("./utility_files/chromedriver_path.txt", 'r') as file:
        for line in file:
            line = line.strip()
            return line
