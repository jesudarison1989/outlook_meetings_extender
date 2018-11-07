from flask import Flask, request, abort
from flask import Flask, render_template
import requests
from flask import jsonify
app = Flask(__name__, template_folder='.')
global accessToken
global userRequester
global userApproval
accessToken = None
userRequester = None
userApproval = None

#--------------------Require Exchange
from pyexchange import Exchange2010Service, ExchangeNTLMAuthConnection
from datetime import datetime
from pytz import timezone
import MySQLdb
from datetime import timedelta

def convert_to_pacific(val1):
    from pytz import timezone
    import pytz
    from datetime import datetime
    date_format='%Y-%m-%d %H:%M:%S %Z'

    val=str(val1)
    val = val.split('+')[0]
    datetime_obj_naive=datetime.strptime(val,"%Y-%m-%d %H:%M:%S")
    #print(datetime_obj_naive)
    #print(type(datetime_obj_naive))
    #print(datetime_obj_naive.tzinfo)
    d = pytz.timezone('utc').localize(datetime_obj_naive)
    #print(d.tzinfo)
    d = d.astimezone(pytz.timezone('US/Pacific'))
    #print(d.tzinfo)
    return d

def mysql_connection(query):
    # Open database connection
    db = MySQLdb.connect("staging-dev","dbowner","0wn3r","ucv_metrics_staging_dev" )

    # prepare a cursor object using cursor() method
    cursor = db.cursor()

    # Prepare SQL query to DELETE required records
    sql = query
    try:
    # Execute the SQL command
        cursor.execute(sql)
        data = cursor.fetchall()
        #print(data)
        # Commit your changes in the database
        db.commit()
    except:
        # Rollback in case there is any error
        db.rollback()
        db.close()

    # disconnect from server
    db.close()

    return data


def list_meetings_and_save_in_DB(room_name):
    ###credentials for outlook exchange:
    URL = u'https://mail.cisco.com/ews/exchange.asmx'
    USERNAME = u'jgerardf'
    PASSWORD = u"Abcd$127"

    # Set up the connection to Exchange
    connection = ExchangeNTLMAuthConnection(url=URL,
    username=USERNAME,
    password=PASSWORD)

    service = Exchange2010Service(connection)

    val = u'{}'.format(room_name)

    print(room_name,val)
    events = service.calendar().list_events(
    start=timezone("US/Pacific").localize(datetime(2018, 11, 8, 7, 0, 0)),
    end=timezone("US/Pacific").localize(datetime(2018, 11, 10, 7, 0, 0)),
    details=True,
    delegate_for=val
    )
    truncate_query = "TRUNCATE TABLE outlook_meetings"

    mysql_connection(truncate_query)

    for events in events.events:
        insert_query = "INSERT INTO outlook_meetings(event_id,event_subject,event_organizer,event_start,event_end,event_start_utc,event_end_utc) VALUES('{}','{}','{}','{}','{}','{}','{}')".format(events.id,events.subject,events.organizer.email,convert_to_pacific(events.start),convert_to_pacific(events.end),events.start,events.end)

        #print(insert_query)

        mysql_connection(insert_query) ##calling to insert meetings info into DB


def findImmediateSuccessorMeetings():

    select_query = "select t1.event_subject as first_subject, t1.event_organizer as first_event_organizer, t1.event_start_utc as first_event_start, t1.event_end_utc as first_event_end, t2.event_subject as second_subject, t2.event_organizer as second_event_organizer, t2.event_start_utc as second_event_start, t2.event_end_utc as second_event_end from outlook_meetings t1 inner join outlook_meetings t2 on (t1.event_end_utc = t2.event_start_utc)"

    data = mysql_connection(select_query)
    data = data[0]

    content = {'first_subject': data[0], 'first_event_organizer': data[1], 'first_event_start': data[2], 'first_event_end': data[3], 'second_subject': data[4], 'second_event_organizer':data[5], 'second_event_start': data[6], 'second_event_end': data[7]}

    return content

def update_meetings(res,start_time,subject,URL,USERNAME,PASSWORD):
    if res == "ok":
        print("ok")
        # Set up the connection to Exchange
        connection = ExchangeNTLMAuthConnection(url=URL,
        username=USERNAME,
        password=PASSWORD)

        service = Exchange2010Service(connection)
        events = service.calendar().list_events(
        start=timezone("US/Pacific").localize(datetime(2018, 11, 8, 7, 0, 0)),
        end=timezone("US/Pacific").localize(datetime(2018, 11, 10, 7, 0, 0)),
        details=True,
        delegate_for=None
        )


        for events in events.events:
            #print(events.start)
            if str(events.start) == start_time:
                print ("------\n Object Type = {type_events}\n start_time = {start}\n end_time = {stop} \n Subject = {subject} \n Organizer = {organiser}\n events_id = {events_id}\n--------------".format(
                start=events.start,
                stop=events.end,
                subject=events.subject,
                organiser = events.organizer.email,
                type_events = type(events),
                events_id = events.id
                ))
                # new_start = events.start+timedelta(minutes=10)
                # events.start = new_start
                # events.update()
                # print('Completed Updating the meeting')
            # else:
            #     #print('time did not match to find the required calendar object')


    else:
        print("Not Updating the meeting")



###Main Program
#### Step :1 ------List all meetings for the provided conference room - Ex: Sony Bono - Alias - CONF_1663@cisco.com
@app.route("/listMeetingsStoreDb/<string:roomName>")
def listMeetingsStoreDb(roomName):
    #room_name = 'CONF_1663@cisco.com' # Sonny Bonno's alias
    list_meetings_and_save_in_DB(roomName)
    return "Successfully inserted into DB!"


### Step :2 ------- Find all the Immediate successor meetings for all the rooms - Ex : Immediate successor meetings for sonny bonno
@app.route("/listSuccessorMeetings")
def listSuccessorMeetings():
    global userRequester
    global userApproval
    #room_name = 'CONF_1663@cisco.com' # Sonny Bonno's alias
    successor_object = findImmediateSuccessorMeetings()
    print(successor_object,type(successor_object))
    #print('First Event Oraganizer details:\n {}\n{}\n{}\n{}\n'.format(successor_object.first_subject,successor_object.first_event_organizer,successor_object.first_event_start,successor_object.first_event_end))
    print('First Event Oraganizer email:\n {}'.format(successor_object['first_event_organizer']))
    print('Second Event Oraganizer email:\n {}'.format(successor_object['second_event_organizer']))

    
    # Registering  to webhook
    registerToWebhook()
    msg = 'Hello Kulwinder, We have noticed that someone have meeting right after you in Sonny Bono. Would you like to extend your meeting.Please respond Yes or No and will take care of the rest'
    userRequester = str(successor_object['first_event_organizer'])
    userApproval = str(successor_object['second_event_organizer'])
    sendUserNotification(msg, userRequester)
    return "Found Successory Meetings"



## Step : 4 ## Update Meetings based on Bot reponse
@app.route("/updateMeetings/<string:res>/<string:username>/<string:password>")
def updateMeetings(res,username,password):
    #room_name = 'CONF_1663@cisco.com' # Sonny Bonno's alias
    URL = u'https://mail.cisco.com/ews/exchange.asmx'
    USERNAME = u'jgerardf'
    PASSWORD = u"Abcd$127"
    start_time = '2018-11-09 20:30:00+00:00'
    subject = 'Test 2 '
    res = 'ok'
    update_meetings(res,start_time,subject,URL,USERNAME,PASSWORD)
    return "Successfully Updated the calendar event!"




### Step : 3 Bot call-----
@app.route("/index")
def index():
    return render_template('index.html', name=index)

@app.route("/index#!/successSubmit")
def success():
    return render_template('success.html', name=success)


@app.route("/requestResp", methods = ["POST"])
def requestResp():
    global userRequester
    action = request.json
    print('action', action["status"])
    if action["status"].lower == "approve":
        msg = "all set, Your request for extend a meeting is Approved"
        sendUserNotification(msg, userRequester)
    else:
        msg = "Sorry your request to extend a meeting is got declined"
        sendUserNotification(msg, userRequester)
    return '', 200

@app.route("/meetingSuggestion", methods = ["POST"])
def meetingSuggestion():
    data = [{
        'room':'room1',
        'biulding':'SJC-12',
        'floor':'3rd'
    },{
        'room':'room2',
        'biulding':'SJC-12',
        'floor':'3rd'
    },{
        'room':'room3',
        'biulding':'SJC-12',
        'floor':'3rd'
    }]
    return jsonify(data)

@app.route("/webhooks", methods = ["POST"])
def webhook():
    if request.method == 'POST':
        print("--------------------------")
        #print(request.json)
        WebhookData = request.json
        data = WebhookData['data']
        msgAccessId = data['id']
        print("msgAccessId ====", msgAccessId)
        if msgAccessId:
            getMessageDetails(msgAccessId);
        return '', 200
    else:
        abort(400)

#---------------------Bot code------------------

#get the msg details of the user
def getMessageDetails(id):
    global userRequester
    global userApproval
    print("-----------getMessageDetails---------------")
    payload = {'Authorization': 'Bearer NTIwZWNiNzMtMTk2Mi00ZDgxLWFkMTQtZDJiNzA4M2VkNDM1ZTg3MzVlMWEtOTc2', 
                'contentType': 'application/json; charset=utf-8'
                }
    r = requests.get("https://api.ciscospark.com/v1/messages/"+id,
                headers=payload
    )
    msg = r.json()
    myMsg = msg['text'].lower()
    #print("userRequester    =====", userRequester)
    #print("myMsg    =====", myMsg)
    #myMsg = 'yes'
    if myMsg == 'yes':
        msg = "Okay, How many minutes would you like to extend eg: 5 , 10"
        sendUserNotification(msg, userRequester)
    if myMsg.isdigit():
        msg = 'Okay great, We will get back to you with decision once your request.'
        numOfMin = myMsg
        sendUserNotification(msg, userRequester)
        msg = 'Hi, We have recived a recived a request regarding your meeting(timing),http://localhost:5000/index'
        sendUserNotification(msg, userApproval)
    return 200
    #print("msg ======>", msg['text'])


#add user to the Bot and send first msg
def sendUserNotification(msg,user):
    print("msg", msg)
    print("user", user)
    payload = {'Authorization': 'Bearer ODA4YzYyNmQtZGZlZS00Yjg2LWI5OTAtZTc0ZDI4N2RiMDAzZDIyZjhhYTktOTE4', 
                'contentType': 'application/json; charset=utf-8'
                }

    body = {
            'text':msg,
            'toPersonEmail':user
    }
    r = requests.post("https://api.ciscospark.com/v1/messages",
                        body,
                        headers=payload
                        )

    print(r.text)

#add user to the Bot and send first msg
def registerToWebhook():
    global accessToken
    if accessToken == None:
        payload = {'Authorization': 'Bearer ODA4YzYyNmQtZGZlZS00Yjg2LWI5OTAtZTc0ZDI4N2RiMDAzZDIyZjhhYTktOTE4', 
                    'contentType': 'application/json; charset=utf-8'
                }

        body = {
                'name':"Your Meetings Assistant",
                'targetUrl':'https://be3a4431.ngrok.io/webhooks'
        }
        r = requests.put("https://api.ciscospark.com/v1/webhooks/Y2lzY29zcGFyazovL3VzL1dFQkhPT0svODk1NGJmZDUtOTM2NC00YjNkLWJiZTgtZmI0M2M3MzNmZmI1",
                            body,
                            headers=payload
                            )
        #global accessToken
        msg = r.json()
        accessToken = msg['id']
        print("accessToken =====>", accessToken)
    else:
        print("Access token already defiend")


#registerToWebhook()
#sendUserNotification("Hello Kulwinder, We have noticed that someone have meeting right after you in Sonny Bono. Would you like to extend your meeting.Please respond Yes or No and will take care of the rest", "kulwisin@cisco.com")

if __name__ == "__main__":
    app.run(debug=True)




# 1 API call : http://localhost:5000/listMeetingsStoreDb/CONF_1663@cisco.com   pushing data to DB
# 2 API call :  http://localhost:5000/listSuccessorMeetings   Finding all succer meeting for friday
# 4 API call: http://localhost:5000/updateMeetings/ok/jgerardf/xxx      Updating the meeting