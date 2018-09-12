from slackclient import SlackClient
import time
import os
import requests
import json
import sys
from datetime import datetime
from datetime import timedelta

"""
This script request Microsoft graph API for calendar then send message through slack api

daily_notify:
Check for today support assignees and send support planning in #support channel
finally it calls weekly_notify
daily_notify is meant to be called each morning from monday to friday

tomorrow_notify:
Check fow tomorrow support assignees and send them personal msg reminder
tomorrow_notify is meant to be called each afternoon from sunday to thursday

weekly_notify
Check for 15 next days not supported time and send warning to #support channel
"""

DEBUG = 1
DEBUG_CHANNEL = ''

CHANNEL = 'support'
SLACK_TOKEN = ''
SLACK_USERNAME = 'Support checker'
SLACK_IMG = 'https://image.flaticon.com/icons/png/512/15/15532.png')
MICRO_CLIENT_ID = ''
MICRO_CLIENT_SECRET = ''
MICRO_TENANT_ID = ''
MICRO_AUTHORITY = 'https://login.microsoftonline.com/'+MICRO_TENANT_ID+'/oauth2/v2.0/token'
MICRO_USER_ID = 'https://login.microsoftonline.com/'+MICRO_TENANT_ID+'/oauth2/v2.0/token'
MICRO_USER_ID = ''

if DEBUG:
  CHANNEL = DEBUG_CHANNEL

people = {
  'TdR': {'name':'', 'id':''},
  'CBU': {'name':'', 'id':''},
  'SHO': {'name':'', 'id':''},
  'MLC': {'name':'', 'id':''},
  'TFU': {'name':'', 'id':''},
  }
#To easily find slack id: 'https://api.slack.com/methods/users.list/test'


print(datetime.now())

def slack_notify(channel, message):
  """Send message to slack channel"""
  sc = SlackClient(SLACK_TOKEN)
  response = sc.api_call(
    "chat.postMessage",
    channel=channel,
    text=message,
    username=SLACK_USERNAME,
    icon_url=SLACK_IMG)
  if response['ok']:
    print('Message posted successfully')
  else:
    print('Error: slack api call failed')
    print(response)

def get_microsoft_api_token():
  post = {'client_id':MICRO_CLIENT_ID, 'scope':'https://graph.microsoft.com/.default', 'client_secret':MICRO_CLIENT_SECRET, 'grant_type':'client_credentials'}
  r = requests.post(MICRO_AUTHORITY, data=post)
  token = json.loads(r.text)['access_token']
  return (token)

def request_api(start, end, subject):
  """Query microsoft graph api to get event whith subject between dates"""
  url = 'https://graph.microsoft.com/v1.0/me/calendarview'
  token = get_microsoft_api_token()
  headers = {
    'Authorization' : 'Bearer {0}'.format(token), 
    'Accept':'application/json', 
    'Prefer': 'outlook.timezone="Europe/Paris"'
  }
  get = {
    'startdatetime': start, 
    'enddatetime': end, 
    '$select':'start,end,subject,categories', 
    '$orderby':'start/dateTime', 
    '$filter':"contains(subject,'"+subject+"')",
    '$top':"1000"
  }
  url = 'https://graph.microsoft.com/v1.0/'+MICRO_TENANT_ID+'/users/'+MICRO_USER_ID+'/calendarview'
  r = requests.get(url, params=get, headers=headers)
  if r.status_code != 200:
    print('Error: microsoft graph api call failed')
    print(r.text)
    sys.exit(0)
  return json.loads(r.text)['value']

def daily_notify():
  """Find today support planning and send it to slack"""
  start = datetime.now().strftime('%Y-%m-%d') + 'T00:00:00'
  end = (datetime.now()+timedelta(days=1)).strftime('%Y-%m-%d') + 'T00:00:00'
  r = request_api(start, end, ' - Support')
  if len(r) > 0:
    msg = ':sunny: Good morning <!channel>, here is the support planning for today:\n'
    for item in r:
      #try to find complete name of support assignee
      who = item['subject'].split()[0]
      who = (people[who]['name'] if who in people.keys() else who)
      msg += ':small_blue_diamond:'+item['start']['dateTime'][11:16]+' -> '+item['end']['dateTime'][11:16]+' : '+who+'\n'
    slack_notify(CHANNEL, msg)
  weekly_notify()

def tomorrow_notify():
  """Find tomorrow support planning and send it to slack to the assignee"""
  start = (datetime.now()+timedelta(days=1)).strftime('%Y-%m-%d') + 'T00:00:00'
  end = (datetime.now()+timedelta(days=2)).strftime('%Y-%m-%d') + 'T00:00:00'
  r = request_api(start, end, ' - Support')
  for item in r:
    #try to find complete name of support assignee
      who = item['subject'].split()[0]
      who = (people[who]['name'] if who in people.keys() else who)
    msg = ':bulb: Hello '+who+', you are on support tomorrow:\n'
    msg += ':small_blue_diamond:'+item['start']['dateTime'][11:16]+' -> '+item['end']['dateTime'][11:16]+'\n'
    if item['subject'].split()[0] in people.keys():
      if DEBUG:
        slack_notify(DEBUG_CHANNEL, msg)
      else: 
        slack_notify(people[item['subject'].split()[0]]['id'], msg)
  
def not_supported(date, r):
  """Return unsupported time range of date in r between 8AM and 6PM
  
  date -- str of the date
  r -- returned by request_api(start, end, ' - Support')
  """
  nosupport = list()
  date = datetime.strptime(date, '%Y-%m-%dT%H:%M:%S')
  #keep only support corresponding to date
  r = list((value) for value in r if value['start']['dateTime'][:10] == datetime.strftime(date, '%Y-%m-%d'))
  for im in range(8 * 60, 18 * 60 + 1):
    check = date + timedelta(minutes=im)
    ok = False
    for item in r:
      if datetime.strptime(item['start']['dateTime'], '%Y-%m-%dT%H:%M:%S.0000000') <= check and datetime.strptime(item['end']['dateTime'], '%Y-%m-%dT%H:%M:%S.0000000') >= check:
        ok = True
    if ok == False:
      nosupport.append(im)
  if len(nosupport) > 0:
    res = range_to_list(nosupport)
    return res
  return False

def range_to_list(times):
  """Return list of time range from single range of minutes"""
  i = 0
  res = list()
  curr = dict()
  curr['start'] = times[0]
  curr['end'] = times[0]
  for minute in times:
    if curr['end'] + 1 == minute:
      curr['end'] = minute
    else:
      res.append(curr)
      curr['start'] = minute
      curr['end'] = minute
  res[-1]['end'] = minute
  for key,item in enumerate(res):
    res[key]['start'] = '{:0>2}:{:0>2}:00'.format(int(item['start']/60), int(item['start']%60))
    res[key]['end'] = '{:0>2}:{:0>2}:00'.format(int(item['end']/60), int(item['end']%60))
  return res
    
def weekly_notify():
  """Look for unsupported time range in next 16 days and send it by slack"""
  start = datetime.now().strftime('%Y-%m-%d') + 'T00:00:00'
  end = (datetime.now()+timedelta(days=16)).strftime('%Y-%m-%d') + 'T00:00:00'
  r = request_api(start, end, ' - Support')
  for id in range(16):
    if (datetime.now()+timedelta(days=id)).strftime('%a') not in ['Sun', 'Sat']:
      day = (datetime.now()+timedelta(days=id)).strftime('%Y-%m-%d') + 'T00:00:00'
      #get all public holidays
      h = list(hol['start']['dateTime'][:10] for hol in request_api(start, end, 'Holiday - ') if 'Yellow Category' in hol['categories'])
      nosupport = not_supported(day, r)
      if nosupport != False and day[:10] not in h:
        msg = ':warning: *Warning*, there is no support on '
        msg += datetime.strptime(day[:10], '%Y-%m-%d').strftime('%A %B %d')
        for item in nosupport:
          msg += '\n:small_red_triangle: from '
          msg += item['start'][:5]
          msg += ' to '
          msg += item['end'][:5]
        slack_notify(CHANNEL, msg)

print(datetime.now())
