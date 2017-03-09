#!python
"""
    DISCLAIMER

    Permission is hereby granted, free of charge, to any person obtaining a
    copy of this software and associated documentation files (the "Software"),
    to deal in the Software without restriction, including without limitation
    the rights to use, copy, modify, merge, publish, distribute, sublicense,
    and/or sell copies of the Software, and to permit persons to whom the
    Software is furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included
    in all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
    THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR
    OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
    ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
    OTHER DEALINGS IN THE SOFTWARE.
"""

import string
import os
import sys
import requests #hTTP GET
from bs4 import BeautifulSoup #Website parser

import icalendar #python-icalendar package
from icalendar import Calendar, Event,vDatetime
import pytz #timezones

import datetime
from dateutil import parser
from dateutil.relativedelta import relativedelta


def parse_Website(cal,url):
    """parses website returns ics"""
    print "Getting URL: " + url
    r = requests.get(url)
    soup = BeautifulSoup(r.text.encode('ascii','ignore'), 'html.parser') #parses website to soup object
    table = soup.find(id="col4_content").find('table') #go to content column and serach for tables
    valid_weekdays = ['MO','DI','MI','DO','FR','SA','SO']
    valid_rowstyles = ["background-color:#dddddd;",'background-color:#ffdddd;'] #table rows with data have no special class
    for table in soup.find_all('table'): #including nested tables
        if(table.get("cellpadding") == "1"): #preselection of day rows
            if(table.get("style") == "font-size:100%; border: none;"): #now we got the right one
                rows = table.findAll('tr') #each row is a day
                for row_counter,tr in enumerate(rows):
                    if(tr.get('style')in valid_rowstyles):
                        cols = tr.findAll('td')
                        if(str(cols[0].text.strip()) in valid_weekdays): #excluding KW Rows
                            cur_weekday =  str(cols[0].text).strip() # from second column
                            cur_date =  str(cols[1].text).strip()
                            cur_todo =  cols[2].findAll('td')
                            print cur_weekday + " " + cur_date
                            if(len(cols)>=2):
                                for sub_col in cur_todo:
                                    subcol_text = sub_col.text.encode('ascii','ignore').strip().replace('\n', ' ').replace('\r', '')
                                    if(sub_col.span is None):
                                        continue
                                    span = sub_col.span.extract()

                                    planer = "unkown"
                                    editor = "unkown"
                                    department = "unkown"

                                    if("Editor:" in subcol_text):
                                        editor = subcol_text[subcol_text.index("Editor:")+7:9999].strip();
                                    if("Planer:" in subcol_text):
                                        planer = subcol_text[subcol_text.index("Planer:")+7:subcol_text.index("Editor:")].strip();
                                    #TODO department = planer[planer.index(";"):99999]
                                    if(editor!="unkown"):
                                        department = editor[editor.index(";")+2:9999].strip()
                                    style = str(sub_col.attrs)
                                    time = str(span.get("title")).strip()
                                    fromto = string.split(time," bis ")
                                    #parsing dates in pyhton to CEST isn't really a thing. For Outlook its mandatory to provide utc timestamps
                                    start_date = pytz.timezone('Europe/Berlin').localize(datetime.datetime.strptime(fromto[0].strip(), '%d.%m.%Y %H:%M'))
                                    start_date = parser.parse(str(start_date.year) + "-" + str(start_date.month) +  "-"+ str(start_date.day) +  " " + str(start_date.hour) +  ":" + str(start_date.minute) +  ":00" + " CEST").astimezone(pytz.utc)
                                    end_date = pytz.timezone('Europe/Berlin').localize(datetime.datetime.strptime(fromto[1].strip(), '%d.%m.%Y %H:%M'))
                                    end_date = parser.parse(str(end_date.year) + "-" + str(end_date.month) +  "-"+ str(end_date.day) +  " " + str(end_date.hour) +  ":" + str(end_date.minute) +  ":00" + " CEST").astimezone(pytz.utc)
                                    title = str(span.text).strip()
                                    title_short = title[:30] + (title[30:] and '..')
                                    print time +" "+ title_short + " "

                                    event = Event()
                                    event.add('summary', title + department) #title of calendar entry
                                    event.add('category', department) #category gets ignored by most clients
                                    event.add('dtstart', start_date)
                                    event.add('description', subcol_text)
                                    event.add('dtend', end_date)
                                    #event.add('X-WR-TIMEZONE', "Europe/Berlin") #confuses outlook 2013
                                    #event.add('X-WR-LOCATION', "Europe/Berlin") #confuses outlook 2013
                                    #create an unique identifier for each event
                                    #Example: 29-04-201818-0030-04-201806-00
                                    event.add('UID', str(fromto[0].strip()+fromto[1]).strip().replace(' ', '').replace('\r', '').replace(':', '-').replace('.', '-'))
                                    cal.add_component(event)
    return cal


def add_months(sourcedate,months):
     """Adding months the complicated way (instead of string replace :)"""
     month = sourcedate.month - 1 + months
     year = int(sourcedate.year + month / 12 )
     month = month % 12 + 1
     day = min(sourcedate.day,calendar.monthrange(year,month)[1])
     return datetime.date(year,month,day)

if __name__ == "__main__":

    #general calendar settings
    cal = Calendar()
    cal.add('prodid', 'fwklux5 ELBE Strahlungsplan p.petring@hzdr.de')
    cal.add('version', '2.0')
    cal.add('x-wr-calname', u"fwklux5 ELBE Strahlungsplan p.petring@hzdr.de")
    cal.add('x-wr-caldesc', u"fwklux5 ELBE Strahlungsplan p.petring@hzdr.de Benutzung auf eigene Gefahr ")
    cal.add('x-wr-relcalid', u"12345")
    cal.add('x-wr-timezone', u"Europe/Berlin")
    cal.add('X-PUBLISHED-TTL', 'PT1H')

    #Timezones are tricky, there can't be to much Timezone-Information
    #also Outlook 2010 requires detailed TZ for non floating events
    tzc = icalendar.Timezone()
    tzc.add('tzid', 'Europe/Berlin')
    tzc.add('x-lic-location', 'Europe/Berlin')

    tzs = icalendar.TimezoneStandard()
    tzs.add('tzname', 'CET')
    tzs.add('dtstart', datetime.datetime(1970, 10, 25, 3, 0, 0))
    tzs.add('rrule', {'freq': 'yearly', 'bymonth': 10, 'byday': '-1su'})
    tzs.add('TZOFFSETFROM', datetime.timedelta(hours=2))
    tzs.add('TZOFFSETTO', datetime.timedelta(hours=1))

    tzd = icalendar.TimezoneDaylight()
    tzd.add('tzname', 'CEST')
    tzd.add('dtstart', datetime.datetime(1970, 3, 29, 2, 0, 0))
    tzs.add('rrule', {'freq': 'yearly', 'bymonth': 3, 'byday': '-1su'})
    tzd.add('TZOFFSETFROM', datetime.timedelta(hours=1))
    tzd.add('TZOFFSETTO', datetime.timedelta(hours=2))

    tzc.add_component(tzs)
    tzc.add_component(tzd)
    cal.add_component(tzc)


    #returning last three months and next 12 moths (+13weeks)
    cur_date=datetime.datetime.now();
    target_date = datetime.datetime.now() + relativedelta(months=12);
    cur_date= cur_date - relativedelta(months=3)

    while (cur_date <= target_date):
        url = 'https://www.hzdr.de/db/!Elbe.StrahlZeitPlan_V2.Liste?pNid=0&pSelAbWoche='+ str(cur_date.day) + '.'+ str(cur_date.month) + '.'+ str(cur_date.year) + '&pSelAnsicht=1'
        try:
            parse_Website(cal,url)
        except Exception,exc:
            #Exit and set Error Code for automated E-Mails if fails
            print str(exc)
            sys.exit("Error while fetching URL: " +url)
        cur_date= cur_date+ relativedelta(months=3)
    #write calendar to disk
    f = open(os.path.join("./", 'strahlzeiten.ics'), 'wb')
    f.write(cal.to_ical())
    f.close()
