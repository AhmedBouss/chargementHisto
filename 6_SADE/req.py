# -*- coding: utf-8 -*-
from datetime import datetime


if 1 :
  from urllib.request import urlopen
  Dep = 43574
  debut = datetime(1900, 1, 1, 0, 0, 0)

  res = urlopen('http://www.python.org')
  headers = res.info()
  date = headers['date'].split(' ')
  date = date[1:4]
  date = " ".join(date)

  newdate = datetime.date(datetime.strptime(str(date), '%d %b %Y'))

  datefin = datetime(newdate.year, newdate.month, newdate.day, 0, 0, 0)

  DPL = (datefin - debut).days + 2
