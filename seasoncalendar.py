
from math import pi, cos


def julianday_springsolstice(year):
    if -1000 < year < 1000:
        y = year / 1000.
        jden = 1721139.29189 + 365242.1374*y + 0.06134*y*y + 0.00111*y*y*y - 0.00071*y*y*y*y
    elif 1000 <= year < 3000:
        y = (year-2000) / 1000.
        jden = 2451623.80984 + 365242.37404*y + 0.05169*y*y - 0.00411*y*y*y - 0.00057*y*y*y*y
    else:
        raise ValueError("Year has to be between -1000 and 3000!")
    return jden


def julianday_summerequinox(year):
    if -1000 < year < 1000:
        y = year / 1000.
        jden = 1721233.25401 + 365241.72562*y - 0.05323*y*y + 0.00907*y*y*y + 0.00025*y*y*y*y
    elif 1000 <= year < 3000:
        y = (year-2000) / 1000.
        jden = 2451716.56767 + 365241.62603*y + 0.00325*y*y + 0.00888*y*y*y - 0.0003*y*y*y*y
    else:
        raise ValueError("Year has to be between -1000 and 3000!")
    return jden


def julianday_autumnsolstice(year):
    if -1000 < year < 1000:
        y = year / 1000.
        jden = 1721325.70455 + 365242.49558*y - 0.11677*y*y - 0.00297*y*y*y + 0.00074*y*y*y*y
    elif 1000 <= year < 3000:
        y = (year-2000) / 1000.
        jden = 2451810.21715 + 365242.01767*y - 0.11575*y*y + 0.00337*y*y*y + 0.00078*y*y*y*y
    else:
        raise ValueError("Year has to be between -1000 and 3000!")
    return jden


def julianday_winterequinox(year):
    if -1000 < year < 1000:
        y = year / 1000.
        jden = 1721414.39987 + 365242.88257*y - 0.00769*y*y - 0.00933*y*y*y - 0.00006*y*y*y*y
    elif 1000 <= year < 3000:
        y = (year-2000) / 1000.
        jden = 2451900.05952 + 365242.74049*y - 0.06223*y*y - 0.00823*y*y*y + 0.00032*y*y*y*y
    else:
        raise ValueError("Year has to be between -1000 and 3000!")
    return jden


S = [
    {'A': 485, 'B': 324.96, 'C': 1934.136},
    {'A': 203, 'B': 337.23, 'C': 32964.467},
    {'A': 199, 'B': 342.08, 'C': 20.186},
    {'A': 182, 'B': 27.85, 'C': 445267.112},
    {'A': 156, 'B': 73.14, 'C': 45036.886},
    {'A': 136, 'B': 171.52, 'C': 22518.443},
    {'A': 77, 'B': 222.54, 'C': 65928.934},
    {'A': 74, 'B': 296.72, 'C': 3034.906},
    {'A': 70, 'B': 243.58, 'C': 9037.513},
    {'A': 58, 'B': 119.81, 'C': 33718.147},
    {'A': 52, 'B': 297.17, 'C': 150.678},
    {'A': 50, 'B': 21.02, 'C': 2281.226},
    {'A': 45, 'B': 247.54, 'C': 29929.562},
    {'A': 44, 'B': 325.15, 'C': 31555.956},
    {'A': 29, 'B': 60.93, 'C': 4443.417},
    {'A': 18, 'B': 155.12, 'C': 67555.328},
    {'A': 17, 'B': 288.79, 'C': 4562.452},
    {'A': 16, 'B': 198.04, 'C': 62894.029},
    {'A': 14, 'B': 199.76, 'C': 31436.921},
    {'A': 12, 'B': 95.39, 'C': 14577.848},
    {'A': 12, 'B': 287.11, 'C': 31931.756},
    {'A': 12, 'B': 320.81, 'C': 34777.259},
    {'A': 9, 'B': 227.73, 'C': 1222.114},
    {'A': 8, 'B': 15.45, 'C': 16859.074}
]


def convert_julianday_to_calendarday(julianday):
    T = (julianday - 2451545) / 36525.
    W = 35999.373*T - 2.47
    Lambda = 1 + 0.0334*cos(W*pi/180.) + 0.0007*cos(2*W*pi/180.)
