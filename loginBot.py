from slackclient import SlackClient
import pprint, openpyxl, os

# TODO - set up enviro variable
slack = SlackClient(open('/Users/ChristianZagazeta/Desktop/token.txt').read())

def sendMessage(message, realName, channel):
    '''
    input: text to send, channel
    output: text to be sent to channel user
    '''
    slack.api_call(
    "chat.postMessage",
    channel=channel,
    text= f"{message}")
    print(f'Message "{message}" sent to {realName}')

def sendMistakeReport(sheet, channel):
    '''
    input: sheet, channel
    output: sends wb[sheet] to user's channel
    '''
    # TODO - remove hardcoding of fileContent
    fileContent = {'file':('PATH.xlsx', open('PATH.xlsx', 'rb'), 'xlsx')}
    slack.api_call(
    'files.upload',
    channels = channel,
    file = fileContent['file'],
    title = 'Mistake Report'
    )
    
def isLogger(memberId):
    '''
    input: memberId
    output: True if not workspace admin, False if workspace admin
    '''
    # admin == workspace admins
    admin = ['USLACKBOT']
    for user in range(len(userList)):
        if userList['members'][user]['is_admin'] == True:
            admin.append(userList['members'][user]['id'])
    if memberId not in admin:
        return True
    return False
        
def users(i):
    '''
    input:  i==userList[i]
            userList contains dict ['members'] that contains a list and dictionaries for each user in workspace
    output: realName = User's registered name
            memberId = User's memberId
            displayName = User's display name
            channel = User's channel (channel is used to message individual users)
    '''
    realName, memberId, displayName = userList['members'][i]['profile']['real_name'], userList['members'][i]['id'], userList['members'][i]['profile']['display_name']
    try:
        channel = slack.api_call('im.open', user = memberId)['channel']['id'] #'im.open' inputs memberId and outputs user's channel under ['channel']['id']                                                                           
    except KeyError: # skips channel
        channel = None
        return realName, memberId, displayName
    return (realName , memberId, channel, displayName)

def hasChannel():
    '''
    input: user
    output: boolean, True if user does have a channel and False if it doesn't
    '''
    if len(user) == 4:
        return True
    return False


# TESTING
userList = slack.api_call('users.list')
for i in range(len(userList)):
    user = users(i)
    if hasChannel(i):
        realName, memberId, channel, displayName = user[0], user[1], user[2], user[3]
        print(realName)
        print(channel)



# -TODO- create mistakereportmockup.xlsx in ZaggyChan

# -TODO- Open workbook, locate 'sheet2'

# -TODO- Begin loop through rows

# -TODO- create new wb, insert name, and mistakes associated with that name

# -TODO- cross reference logger name and userList

# -TODO- if logger isn't in userList, add to list of loggers without slack data
#        move wb with unlisted loggers report to a new folder

# -TODO- prompt user to send mistake report (optional)

# -TODO- save, create log and send new wb to user with channel

# -TODO- delete sent wb (optional)

# -TODO- continue until rows exhausted

# -TODO- Add unlisted loggers memberId or displaynames

# to be continued...