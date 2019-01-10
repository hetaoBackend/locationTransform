# -*- coding: utf-8 -*-
import xlrd
import xlwt
from xlutils.copy import copy
from math import sqrt, sin, cos, atan2, pi

x_pi = pi * 3000.0 / 180.0
ee = 1/298.257223563
a = 6378137.0

def bd09togcj02(bd_lon, bd_lat):
    x = bd_lon - 0.0065
    y = bd_lat - 0.006
    z = sqrt(x**2 + y**2) - 0.00002 * sin(y * x_pi)
    theta = atan2(y, x) - 0.000003 * cos(x * x_pi)
    gg_lng = z * cos(theta)
    gg_lat = z * sin(theta)
    return gg_lng, gg_lat

def transformlat(lng, lat):
	ret = -100.0 + 2.0 * lng + 3.0 * lat + 0.2 * lat * lat + 0.1 * lng * lat + 0.2 * sqrt(abs(lng))
	ret += (20.0 * sin(6.0 * lng * pi) + 20.0 * sin(2.0 * lng * pi)) * 2.0 / 3.0
	ret += (20.0 * sin(lat * pi) + 40.0 * sin(lat / 3.0 * pi)) * 2.0 / 3.0
	ret += (160.0 * sin(lat / 12.0 * pi) + 320 * sin(lat * pi / 30.0)) * 2.0 / 3.0
	return ret

def transformlng(lng, lat):
    ret = 300.0 + lng + 2.0 * lat + 0.1 * lng * lng + 0.1 * lng * lat + 0.1 * sqrt(abs(lng))
    ret += (20.0 * sin(6.0 * lng * pi) + 20.0 * sin(2.0 * lng * pi)) * 2.0 / 3.0
    ret += (20.0 * sin(lng * pi) + 40.0 * sin(lng / 3.0 * pi)) * 2.0 / 3.0
    ret += (150.0 * sin(lng / 12.0 * pi) + 300.0 * sin(lng / 30.0 * pi)) * 2.0 / 3.0
    return ret

def gcj02towgs84(lng, lat):
    dlat = transformlat(lng - 105.0, lat - 35.0)
    dlng = transformlng(lng - 105.0, lat - 35.0)
    radlat = lat / 180.0 * pi
    magic = sin(radlat)
    magic = 1 - ee * magic * magic
    sqrtmagic = sqrt(magic)
    dlat = (dlat * 180.0) / ((a * (1 - ee)) / (magic * sqrtmagic) * pi)
    dlng = (dlng * 180.0) / (a / sqrtmagic * cos(radlat) * pi)
    mglat = lat + dlat
    mglng = lng + dlng
    return lng * 2 - mglng, lat * 2 - mglat

def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))

if __name__ == "__main__":
    # process the school table
    data = open_excel('东城小学坐标.xlsx')
    excel = xlwt.Workbook()
    rdTable = data.sheets()[0]
    wtTable = excel.add_sheet("Sheet1", cell_overwrite_ok=True)
    nrows = rdTable.nrows
    wtTable.write(0,0,rdTable.row_values(0)[0])
    wtTable.write(0,1,rdTable.row_values(0)[1])
    wtTable.write(0,2,"WGS84_X")
    wtTable.write(0,3,"WGS84_Y")
    for i in range(1, nrows):
        bd_lon, bd_lat = float(rdTable.row_values(i)[2]), float(rdTable.row_values(i)[3])
        bd_lon, bd_lat = bd09togcj02(bd_lon, bd_lat)
        wgs84_lon, wgs84_lat = gcj02towgs84(bd_lon, bd_lat)
        wtTable.write(i,0,rdTable.row_values(i)[0])
        wtTable.write(i,1,rdTable.row_values(i)[1])
        wtTable.write(i,2,str(wgs84_lon))
        wtTable.write(i,3,str(wgs84_lat))
    excel.save('eastCitySchool_WGS84.xls')

    # process the eastCity table
    data = open_excel('东城坐标.xlsx')
    excel = xlwt.Workbook()
    rdTable = data.sheets()[0]
    wtTable = excel.add_sheet("Sheet1", cell_overwrite_ok=True)
    nrows = rdTable.nrows
    wtTable.write(0,0,rdTable.row_values(0)[0])
    wtTable.write(0,1,"WGS84_X")
    wtTable.write(0,1,"WGS84_Y")
    for i in range(1, nrows):
        bd_lon, bd_lat = float(rdTable.row_values(i)[1]), float(rdTable.row_values(i)[2])
        bd_lon, bd_lat = bd09togcj02(bd_lon, bd_lat)
        wgs84_lon, wgs84_lat = gcj02towgs84(bd_lon, bd_lat)
        wtTable.write(i,0,rdTable.row_values(i)[0])
        wtTable.write(i,1,str(wgs84_lon))
        wtTable.write(i,2,str(wgs84_lat))
    excel.save('eastCity_WGS84.xls')
