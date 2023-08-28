import datetime
from astral.geocoder import database, lookup
from astral.moon import moonrise, moonset

def test():
  format_data = "%m/%d/%y %H:%M:%S.%f %Z"

  print("====== test missing data on 5/17/23 =========")
  day = datetime.date(year=2023, month=5, day=17)

  slc = lookup("Salt Lake City", database())
  try:
    print(moonset(slc.observer, day).strftime(format_data))
  except ValueError as e:
    print(f'caught exception: {str(e)}')

test()
