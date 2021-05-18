import datetime
import calendar
from datetime import tzinfo, timezone

import time


dt = datetime.datetime.now()
print('dt', dt)
#dt = dt.tzinfo
#print(dt)
print('timetz               ', dt.timetz())
print('utcoffset            ', dt.utcoffset())
print('astimezone           ', dt.astimezone())
print('tzinfo               ', dt.tzinfo)
print('dst                  ', dt.dst())
print('toordinal            ', dt.toordinal())
print('fromordinal(32)      ', dt.fromordinal(32))
print('utcnow               ', dt.utcnow())
print('timestamp            ', dt.timestamp())
print('dt.utcfromtimestamp(dt.timestamp())', dt.utcfromtimestamp(dt.timestamp()))
print('tzinfo               ', dt.tzinfo)
print('time.timezone        ', time.timezone)
print('time.daylight        ', time.daylight)
print('time.tzname[0]       ', time.tzname[0])
print('time.altzone         ', time.altzone)
print('time.localtime       ', time.localtime(1611747319.481842)[:5])
print('time.asctime()       ', time.asctime())
print('dt.utctimetuple()[:5]', dt.utctimetuple()[:5])
print('timedelta            ', datetime.timedelta.min)
print(time.tzname)
print(timezone.tzname(datetime.datetime.now()))

