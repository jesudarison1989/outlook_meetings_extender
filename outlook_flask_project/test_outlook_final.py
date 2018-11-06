from pyexchange import Exchange2010Service, ExchangeNTLMAuthConnection
from datetime import datetime
from pytz import timezone
import MySQLdb
from datetime import timedelta

##sudo pip show pyexchange | grep Version -> To check module version


###credentials for outlook exchange:
URL = u'https://mail.cisco.com/ews/exchange.asmx'
USERNAME = u'jgerardf'
PASSWORD = u"Abcd$127"

#USERNAME = u'CONF_1663'
#PASSWORD = u"C!sc0123"

# Set up the connection to Exchange 
connection = ExchangeNTLMAuthConnection(url=URL,
                                        username=USERNAME,
                                        password=PASSWORD)

service = Exchange2010Service(connection)


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

# Creating new meetings:

def create_new_meeting():
    # You can set event properties when you instantiate the event...
    event = service.calendar().new_event(
    subject=u"80s Movie Night",
    attendees=[u'jgerardf@cisco.com', u'kulwisin@cisco.com'],
    location = u"My house",
    )

    # ...or afterwards
    event.start=timezone("US/Pacific").localize(datetime(2018,11,5,17,0,0))
    event.end=timezone("US/Pacific").localize(datetime(2018,11,5,18,0,0))

    event.html_body = u"""<html>
        <body>
            <h1>80s Movie night</h1>
            <p>We're watching Spaceballs, Wayne's World, and
            Bill and Ted's Excellent Adventure.</p>
            <p>PARTY ON DUDES!</p>
        </body>
    </html>"""

    # Connect to Exchange and create the event
    event.create()


def list_meetings_and_save_in_DB():  

    events = service.calendar().list_events(
        start=timezone("US/Pacific").localize(datetime(2018, 11, 5, 7, 0, 0)),
        end=timezone("US/Pacific").localize(datetime(2018, 11, 7, 4, 0, 0)),
        details=True,
        delegate_for=u'CONF_1663@cisco.com'
    )
    truncate_query = "TRUNCATE TABLE outlook_meetings"

    mysql_connection(truncate_query)

    for events in events.events:
        # print ("------\n Object Type = {type_events}\n start_time = {start}\n end_time = {stop} \n Subject = {subject} \n Organizer = {organiser}\n events_id = {events_id}\n--------------".format(
        #     start=events.start,
        #     stop=events.end,
        #     subject=events.subject,
        #     organiser = events.organizer.email,
        #     type_events = type(events),
        #     events_id = events.id
        # ))
        #truncate_query = "TRUNCATE TABLE outlook_meetings"

        #mysql_connection(truncate_query)

        insert_query = "INSERT INTO outlook_meetings(event_id,event_subject,event_organizer,event_start,event_end,event_start_utc,event_end_utc) VALUES('{}','{}','{}','{}','{}','{}','{}')".format(events.id,events.subject,events.organizer.email,convert_to_pacific(events.start),convert_to_pacific(events.end),events.start,events.end)
        
        print(insert_query)
        
        mysql_connection(insert_query)  ##calling to insert meetings info into DB
        

def find_recurrent_meetings():
    #select_query = "select t1.event_id,t2.event_id,t1.event_start as first_event_start,t1.event_end as first_event_end,t2.event_start as second_event_start,t2.event_end from outlook_meetings t1 inner join outlook_meetings t2 on (t1.event_end = t2.event_start )"

    select_query = "select second_subject,second_event_start from (select t1.event_subject as first_subject,t2.event_subject as second_subject,t1.event_start_utc as first_event_start,t1.event_end_utc as first_event_end,t2.event_start_utc as second_event_start,t2.event_end_utc from outlook_meetings t1 inner join outlook_meetings t2 on (t1.event_end_utc = t2.event_start_utc ))a"
    
    data = mysql_connection(select_query)
    return data

def update_meetings(res,start_time,subject,URL,USERNAME,PASSWORD):
    if res == "ok":
        print("ok")
        # Set up the connection to Exchange 
        connection = ExchangeNTLMAuthConnection(url=URL,
                                        username=USERNAME,
                                        password=PASSWORD)

        service = Exchange2010Service(connection)
        events = service.calendar().list_events(
        start=timezone("US/Pacific").localize(datetime(2018, 11, 5, 7, 0, 0)),
        end=timezone("US/Pacific").localize(datetime(2018, 11, 7, 4, 0, 0)),
        details=True,
        delegate_for=None
        )
    
        for events in events.events:
            if str(events.start) == '2018-11-06 00:00:00+00:00':
                # print ("------\n Object Type = {type_events}\n start_time = {start}\n end_time = {stop} \n Subject = {subject} \n Organizer = {organiser}\n events_id = {events_id}\n--------------".format(
                #     start=events.start,
                #     stop=events.end,
                #     subject=events.subject,
                #     organiser = events.organizer.email,
                #     type_events = type(events),
                #     events_id = events.id
                # ))
                new_start =  events.start+timedelta(minutes=10)
                events.start = new_start
                events.update()
                print('Completed Updating the meeting')
        

    else:
        print("Not Updating the meeting")

        
def mysql_ctbx_connection(query):  
    # Open database connection
    db = MySQLdb.connect("ctbx-mysql-prod-a-vip.cisco.com","readonly","readonly","ctbx_portal" )

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
       
### Main program starts here: 
## Step 1:       
list_meetings_and_save_in_DB()
#
##Step :2  
#data = find_recurrent_meetings()
#print(data)

#Step: 3 Alert the BOT, send messages and get a response.
URL = u'https://mail.cisco.com/ews/exchange.asmx'
USERNAME = u'jgerardf'
PASSWORD = u"Abcd$127"
start_time = '2018-11-06 00:00:00+00:00'
subject = 'Test meeting 4'
update_meetings('ok',start_time,subject,URL,USERNAME,PASSWORD)

#Step: 4 Update the meeting based on response
#recurring_list = data 
#update_meetings('ok','n/a','')


#Step: 5 Get other room options to book if available:
# query = "select full_name from wpr_conference_room where  building_name='SJ-12' and floor_id = 3 and room_type = 'Collaborative Technology Room' and full_name  not like '%SONNY%'"
# data = mysql_ctbx_connection(query)
# print('----\n\n {}------------------ '.format(data))





       


