import boto3
import botocore.exceptions
from scripts.utils import get_aws_credentials
from decimal import Decimal


def upload_plays(data, date):
    creds = get_aws_credentials()
    stringified_data = {}
    for num, game in data.items():
        num_str = str(num)
        for key, stat in game['away_team'].items():
            if isinstance(stat, float):
                game['away_team'][key] = Decimal(str(stat))
        for key, stat in game['home_team'].items():
            if isinstance(stat, float):
                game['home_team'][key] = Decimal(str(stat))
        game['spread'] = Decimal(str(game['spread']))
        stringified_data[num_str] = game

    dynamodb = boto3.resource(
        'dynamodb',
        region_name=creds['region_name'],
        aws_access_key_id=creds['access_key'],
        aws_secret_access_key=creds['secret_key']
    )
    table = dynamodb.Table('cbb_system')
    json_data = {'data': stringified_data, 'date': date}
    try:
        table.put_item(Item=json_data)
    except botocore.exceptions.ClientError as error:
        print(f"Error code: {error.response['Error']['Code']}, Error message: {
              error.response['Error']['Message']}")
    except Exception as error:
        print(f"{error}")


def get_plays(date):
    creds = get_aws_credentials()
    dynamodb = boto3.resource(
        'dynamodb',
        region_name=creds['region_name'],
        aws_access_key_id=creds['access_key'],
        aws_secret_access_key=creds['secret_key']
    )
    table = dynamodb.Table('cbb_system')
    try:
        response = table.get_item(Key={'date': date}).get('Item', None)
        return response
    except:
        print(f"Error Fetching {date} from Dynamo")
