from flask import Flask
import requests
from pyexchange import Exchange2010Service, ExchangeNTLMAuthConnection
from datetime import datetime
from pytz import timezone
import MySQLdb
from datetime import timedelta
from flask import jsonify


app = Flask(__name__)
 
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

def list_meetings_and_save_in_DB(rooms_list):  
    ###credentials for outlook exchange:
    URL = u'https://mail.cisco.com/ews/exchange.asmx'
    USERNAME = u'jgerardf'
    PASSWORD = u"Abcd$127"

    # Set up the connection to Exchange 
    connection = ExchangeNTLMAuthConnection(url=URL,
                                            username=USERNAME,
                                            password=PASSWORD)

    service = Exchange2010Service(connection)
    
    truncate_query = "TRUNCATE TABLE outlook_meetings"

    mysql_connection(truncate_query)

    for room_name in rooms_list.keys():
        room_name = rooms_list[room_name]
        val = u'{}'.format(room_name)

        print(room_name,val)
        events = service.calendar().list_events(
            start=timezone("US/Pacific").localize(datetime(2018, 11, 8, 7, 0, 0)),
            end=timezone("US/Pacific").localize(datetime(2018, 11, 10, 7, 0, 0)),
            details=True,
            delegate_for=val
        )
        

        for events in events.events:
            insert_query = "INSERT INTO outlook_meetings(event_id,event_organizer_name,event_subject,event_organizer,event_start,event_end,event_start_utc,event_end_utc,event_location) VALUES('{}','{}','{}','{}','{}','{}','{}','{}','{}')".format(events.id,events.organizer.name,events.subject,events.organizer.email,convert_to_pacific(events.start),convert_to_pacific(events.end),events.start,events.end,events.location)
            
            #print(insert_query)
            
            mysql_connection(insert_query)  ##calling to insert meetings info into DB


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

        #print(events.events)
    
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
                # new_start =  events.start+timedelta(minutes=10)
                # events.start = new_start
                # events.update()
                # print('Completed Updating the meeting')
            else:
                print('time did not match to find the required calendar object')
        

    else:
        print("Not Updating the meeting")


def get_another_room():

    ### The start date time in utc and end date time in utc is obtained from the previous api which is providing the immediate successor meeting
    select_query = "select distinct(room_name) from (select *, case when event_location like '%SONNY BONO%' then 'sonny_bonno' when event_location like '%BOB MARLEY%' then 'bob_marley' when event_location like '%MADONNA%' then 'madonna' when event_location like '%VAN MORRISON%' then 'van_morrisson' end as 'room_name' from outlook_meetings where event_start<='2018-11-09 13:00:00-08:00' and `event_end` >='2018-11-09 12:30:00-08:00' )a;"
        
    data = mysql_connection(select_query)
   
    set_a_available_rooms = {'sonny_bonno','bob_marley','madonna','van_morrison'}
    set_b_not_available_rooms = set()
    
    for room in data: 
        set_b_not_available_rooms.add(room[0])

    rooms_available = set_a_available_rooms-set_b_not_available_rooms
     
    return rooms_available

    # content = {'first_subject': data[0], 'first_event_organizer': data[1], 'first_event_start': data[2], 'first_event_end': data[3]}





##  Routes declarations :
@app.route("/")
def index():
    #list_meetings_and_save_in_DB()
    return "Welome to App!"
 

#### Step :1  ------List all meetings for the provided conference room - Ex: Sony Bono - Alias -  CONF_1663@cisco.com
@app.route("/listMeetingsStoreDb/<string:roomName>")
def listMeetingsStoreDb(roomName):
    
    room_obj = {'sonny_bonno':'CONF_1663@cisco.com','bob_marley':'CONF_2996@cisco.com','madonna':'CONF_69814@cisco.com','van_morrison':'CONF_1981@cisco.com'}

    list_meetings_and_save_in_DB(room_obj)
    return "Successfully inserted into DB!"
    


### Step :2 ------- Find all the Immediate successor meetings for all the rooms  - Ex : Immediate successor meetings  for sonny bonno
@app.route("/listSuccessorMeetings")
def listSuccessorMeetings():
    #room_name = 'CONF_1663@cisco.com' # Sonny Bonno's alias
    successor_object = findImmediateSuccessorMeetings()
    print(successor_object,type(successor_object))
    #print('First Event Oraganizer details:\n {}\n{}\n{}\n{}\n'.format(successor_object.first_subject,successor_object.first_event_organizer,successor_object.first_event_start,successor_object.first_event_end))
    print('First Event Oraganizer email:\n {}'.format(successor_object['first_event_organizer']))
    print('Second Event Oraganizer email:\n {}'.format(successor_object['second_event_organizer']))
    return "Found Successory Meetings"
 

#-------------------------------------------------------------------------------
### Step : 3---> BOT CODE

@app.route("/webhooks")
def webhooks():
    return "webhooks"

# @app.route("/members/<string:name>/")
# def getMember(name):
#     name = name+'----*******'
#     return name

#----------------------------------------------------------------------------------


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


### step : 6 ## Get another confernce room:
@app.route("/getAnotherRoom")
def getAnotherRoom():
    rooms_available = get_another_room()
    print(rooms_available)
    return jsonify(','.join(list(rooms_available)))





if __name__ == "__main__":
    app.run(debug=True)
