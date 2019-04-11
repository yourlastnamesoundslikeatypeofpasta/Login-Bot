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
    # TODO - remove hardcoding of fileContent
    fileContent = {'file':('/Users/ChristianZagazeta/Desktop/Scripts/Mistake Counter/mistakereport.xlsx', open('/Users/ChristianZagazeta/Desktop/Scripts/Mistake Counter/mistakereport.xlsx', 'rb'), 'xlsx')}
    slack.api_call(
    'files.upload',
    channels = channel,
    file = fileContent['file'],
    title = 'Mistake Report'
    )
    
def isLogger(memberId):
    # admin == workspace admins
    admin = ['USLACKBOT']
    for user in range(len(userList)):
        if userList['members'][user]['is_admin'] == True:
            admin.append(userList['members'][user]['id'])
    if memberId not in admin:
        return True
    return False
        
def users(i):
    realName, memberId, displayName = userList['members'][i]['profile']['real_name'], userList['members'][i]['id'], userList['members'][i]['profile']['display_name']
    try:
        channel = slack.api_call('im.open', user = memberId)['channel']['id']
    except KeyError:
        channel = None
        print(f"{realName} does not have a channel listed: skipped")
        print(f"Real Name: {realName}", f"Member ID: {memberId}")
        return realName, memberId, displayName
    print(f"Real Name: {realName}", f"Member ID: {memberId}", f'Channel: {channel}')  
    return (realName , memberId, channel, displayName)

def hasChannel(i):
    if len(user) == 4:
        return True
    return False


# TODO - 
userList = slack.api_call('users.list')
for i in range(len(userList)):
    user = users(i)
    if hasChannel(i):
        realName, memberId, channel, displayName = user[0], user[1], user[2], user[3]
        print(realName)