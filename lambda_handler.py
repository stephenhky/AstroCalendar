import json

from seasoncalendar import *

def lambda_handler(event, context):
    print(event)
    if 'year' in event and 'season' in event:
        year = event['year']
        season = event['season']
    else:
        body = json.loads(event['body'])
        year = body['year']
        season = body['season']

    if season.lower() == 'spring':
        julianday = julianday_springsolstice(year)
    elif season.lower() == 'summer':
        julianday = julianday_summerequinox(year)
    elif season.lower() == 'autumn':
        julianday = julianday_autumnsolstice(year)
    elif season.lower() == 'winter':
        julianday = julianday_winterequinox(year)
    else:
        return {
            'statusCode': 501,
            'body': json.dumps('No season called {}'.format(season))
        }

    t = convert_julianday_to_calendarday(julianday)
    body = {
        'year': year,
        'season': season,
        'julianday': julianday,
        'month': t.month,
        'day': t.day,
        'hour': t.hour,
        'minute': t.minute,
        'second': t.second
    }

    return {
        'statusCode': 200,
        'body': json.dumps(body)
    }
