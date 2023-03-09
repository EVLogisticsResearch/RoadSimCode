import pypyodbc
import adodbapi  # (from pywin32)
import xlrd
import xlwt
import numpy as npa
import matplotlib.pyplot as plt
import pandas as pd 
from sqlalchemy import create_engine
import urllib.request
import json
import polyline
import pprint
#import pyodbc
import urllib
import math  
import random
import geopy.distance
import re
import os
import csv
import copy
num_regex = re.compile(r'[0-9]')
char_regex = re.compile(r'[a-zA-z]')
Google_API_Key='AIzaSyDp0eSEkc7XoODRShW5KAqN7GXIoBbTM4M'

Set_CFer=0.7
Set_VSlopeP=0.5031*Set_CFer
Set_VSlopeN=0.4367*Set_CFer
Max_Speed_Cruise=5.55        #20kph
Max_Speed_Collection=3.5    #12.6kph
Collection_Work_Time=10     #in second
Stop_Start_Distance=35
TL_Delay_Max=38

PostcodeAtColumn = 1#Starting at 0
DomAddrAtColumn = 2
RecAddrAtColumn = 5
DomGroupAtColumn = 8
RecGroupAtColumn = 9
LatitudeAtColumn = 10
LongitudeAtColumn = 11

DB_AddressatColumn= 1#Starting at 0
DB_LatitudeatColumn=3
DB_LongitudeatColumn=4
DB_AltitudeatColumn=5

DB_TypeatColumn=1
DB_Distancee_m_atColumn=6
DB_SlopeatColumn=4

CSV_LatitudeatColumn=3# order in TEMP.XLSX NOT in CSV so csv position+1
CSV_LongitudeatColumn=4
CSV_AccuDistanceatColumn=6
CSV_DistanceatColumn=7

Min_Form_Lenth=6#minimum form length

Bin_file = 'Sheffield_Bins_Map_Full.xlsx'
CSV_path_Add_Distance = "C:\\Python\\RoadSim\\toAddDistance\\" 
DomGroupName=[]


Depot_Lat=53.386837
Depot_Long=-1.448617

def addtodict3(thedict,key_a,key_b,key_c,val):
    if key_a in thedict:
        if key_b in thedict[key_a]:
            thedict[key_a][key_b].update({key_c:val})
        else:
            thedict[key_a].update({key_b:{key_c:val}})
    else:
        thedict.update({key_a:{key_b:{key_c:val}}})

def SpeedProfileBuilder():
    conn = pypyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\\Python\\RoadSimDom.mdb;")
    Crsr_Form = conn.cursor()
    Crsr_Contant = conn.cursor()
    Crsr_Elev = conn.cursor()  
    SQL = "SELECT name FROM MSYSOBJECTS WHERE ((TYPE=1) and flags=0 AND (name NOT LIKE '*MSys*'))"
    SQL_Rtn=Crsr_Form.execute(SQL).fetchall()

    connection_string = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Python\RoadSimDom_Route.mdb;'
    r'ExtendedAnsiSQL=1;'
    )
    connection_uri = f"access+pyodbc:///?odbc_connect={urllib.parse.quote_plus(connection_string)}"
    engine = create_engine(connection_uri)

    for Forms in SQL_Rtn:
        OutType=[]
        OutTim=[]
        OutSlope=[]
        OutSpeed=[]
        
        print(str(Forms))
        if str(Forms) == "('Misc ',)" or str(Forms) == "(': ',)":
            continue
        Form_Name=str(Forms).replace("'","").replace("(","").replace(")","").replace(",","")
        SQL_Form=("select * from [" + Form_Name +"]")
        Form_Cont = Crsr_Contant.execute (SQL_Form).fetchall()

        _F_Lenth=len(Form_Cont)
        if _F_Lenth<Min_Form_Lenth:
            continue

        while Form_Cont[_i][DB_TypeatColumn]!='N':#Jmp out the leading collection section
            _i=_i+1
        Start_row=_i

        Section_Distance=0
        Section_SlopDistance=[]
        Section_SlopList=[]
        Sim_Time=0
        while _i<_F_Lenth-1:
            _TCurr=Form_Cont[_i][DB_TypeatColumn]
            _TNext=Form_Cont[_i+1][DB_TypeatColumn]
            #Cruise section accumulator 
            if _TCurr=='N' and _TNext=='N':
                Section_Distance=Section_Distance+(float)(Form_Cont[_i][DB_Distancee_m_atColumn])
                if Form_Cont[_i][DB_SlopeatColumn]=="":
                    Section_SlopList.append(0)
                else:
                    Section_SlopList.append((float)(Form_Cont[_i][DB_SlopeatColumn]))
                Section_SlopDistance.append(Section_Distance)
                _i=_i+1
            #Cruise section simulator
            if _TCurr=='N' and _TNext=='C':# This indicates the end of current cruise section
                Section_Distance=Section_Distance+Form_Cont[_i][DB_Distancee_m_atColumn] #add the last cell
                #Section_SlopList.append(Section_SlopList[len(Section_SlopList)-1])
                if Form_Cont[_i][DB_SlopeatColumn]=="":
                    Section_SlopList.append(0)
                else:
                    Section_SlopList.append((float)(Form_Cont[_i][DB_SlopeatColumn]))
                _i=_i+1
                Sim_Speed=0
                Sim_Distance=0
                Section_Sim_Time=0
                _i_Slope=0
                Section_SlopDistance.append(9999)#Used from last SlopDistance value to section end
                #Sub simulator for acceleration and const speed section -> /---- 
                while Sim_Distance<Section_Distance-(Max_Speed_Cruise*Max_Speed_Cruise)/(Set_VSlopeN*2):
                    OutType.append('N')
                    Sim_Time=Sim_Time+1
                    OutTim.append(Sim_Time)

                    if Sim_Distance<Section_SlopDistance[_i_Slope]:
                        OutSlope.append(Section_SlopList[_i_Slope])
                    else:
                        _i_Slope=_i_Slope+1
                        OutSlope.append(Section_SlopList[_i_Slope])

                    if (Section_Sim_Time*Set_VSlopeP>Max_Speed_Cruise):
                        Current_Speed=Max_Speed_Cruise
                        Sim_Distance=Sim_Distance+Max_Speed_Cruise
                    else:
                        Current_Speed=Section_Sim_Time*Set_VSlopeP
                        Sim_Distance=Section_Sim_Time*Set_VSlopeP/2
                    OutSpeed.append(Current_Speed)
                    Section_Sim_Time=Section_Sim_Time+1          
                #Rst Section_Sim_Time so it becoming the pointer for deceleration simulator 
                Section_Sim_Time=0
                #Sub simulator for deceleration section -> \
                while Sim_Distance<Section_Distance:
                    OutType.append('N')
                    Sim_Time=Sim_Time+1
                    OutTim.append(Sim_Time)

                    if Sim_Distance<Section_SlopDistance[_i_Slope]:
                        OutSlope.append(Section_SlopList[_i_Slope])
                    else:
                        _i_Slope=_i_Slope+1
                        OutSlope.append(Section_SlopList[_i_Slope])

                    Section_Sim_Time=Section_Sim_Time+1
                    OutSpeed.append(Max_Speed_Cruise-Section_Sim_Time*Set_VSlopeN)
                    Sim_Distance=Sim_Distance+abs(Max_Speed_Cruise-Section_Sim_Time*Set_VSlopeN+Set_VSlopeN/2)#PPOE
                #Use following to add a '0' at the end of each section
                #OutType.append('N')
                #Sim_Time=Sim_Time+1
                #OutTim.append(Sim_Time)
                #OutSpeed.append(0)
                Section_Distance=0
                Section_SlopDistance.clear()
                Section_SlopList.clear()
            #Collection section accumulator
            #Cycle termination theory: 1. D>Dmin 2. D+Dnext<Dmax 3. Tcurr<>Tnext
            if _TCurr=='C' and _TNext=='C':#Collection section 
                Section_Distance=Section_Distance+Form_Cont[_i][DB_Distancee_m_atColumn]
                if Form_Cont[_i][DB_SlopeatColumn]=="":
                    Section_SlopList.append(0)
                else:
                    Section_SlopList.append((float)(Form_Cont[_i][DB_SlopeatColumn]))
                Section_SlopDistance.append(Section_Distance)
                _i=_i+1
            #Collection section simulator
            if _TCurr=='C' and _TNext=='N':
                if Form_Cont[_i-1][DB_TypeatColumn]=='N':#traffic light stop#################################
                    TL_Tim=0
                    _i=_i+1
                    TL_Delay=random.randint(0,TL_Delay_Max)
                    while TL_Tim<=TL_Delay:
                        OutType.append('P')
                        Sim_Time=Sim_Time+1
                        OutTim.append(Sim_Time)

                        OutSlope.append(OutSlope[len(OutSlope)-1])

                        OutSpeed.append(0)
                        TL_Tim=TL_Tim+1
                else:
                    Section_Distance=Section_Distance+Form_Cont[_i][DB_Distancee_m_atColumn] #add the last cell
                    #Section_SlopList.append(Section_SlopList[len(Section_SlopList)-1])
                    if Form_Cont[_i][DB_SlopeatColumn]=="":
                        Section_SlopList.append(0)
                    else:
                        Section_SlopList.append((float)(Form_Cont[_i][DB_SlopeatColumn]))
                    _i=_i+1
                    _i_Slope=0        
                    Sim_Distance=0
                    Section_SlopDistance.append(9999)#Used from last SlopDistance value to section end

                    #Stop_Start section speed if using a triangle approximation
                    Stop_Start_Speed=math.sqrt((2*Set_VSlopeP*Set_VSlopeN*Stop_Start_Distance)/(Set_VSlopeP+Set_VSlopeN))
                    if Stop_Start_Speed < Max_Speed_Collection:
                        #1. cannot reach maximum speed
                        Sim_Section_Max_Speed=Stop_Start_Speed
                    else:
                        #2. maximum speed reached
                        Sim_Section_Max_Speed=Max_Speed_Collection
                    Sim_Overall_Section_Distance=0
                    #########################################################################DDDDDDDDOOOOOOOOWWWWWNNNNN
                    #Calculate each section before the remaining section<Stop_Start_Distance
                    while Section_Distance-Sim_Overall_Section_Distance>Stop_Start_Distance:
                        Sim_Speed=0
                        Stop_Start_Sim_Distance=0
                        Stop_Start_Sim_Time=0

                        #Stop start section acceleration part simulator
                        while Sim_Speed<=Sim_Section_Max_Speed:#Acceleration
                            OutType.append('C')
                            Sim_Time=Sim_Time+1
                            OutTim.append(Sim_Time)

                            if Sim_Distance<Section_SlopDistance[_i_Slope]:
                                OutSlope.append(Section_SlopList[_i_Slope])
                            else:
                                _i_Slope=_i_Slope+1
                                OutSlope.append(Section_SlopList[_i_Slope])

                            Sim_Speed=Stop_Start_Sim_Time*Set_VSlopeP
                            OutSpeed.append(Sim_Speed)
                            Stop_Start_Sim_Time=Stop_Start_Sim_Time+1
                            Stop_Start_Sim_Distance=Stop_Start_Sim_Distance+(Sim_Speed-0.5*Set_VSlopeP)
                            Sim_Distance=Sim_Distance+(Sim_Speed-0.5*Set_VSlopeP)
                        #Stop start section const-speed part simulator  
                        while Stop_Start_Sim_Distance<Stop_Start_Distance-(0.5*Sim_Section_Max_Speed*Sim_Section_Max_Speed/Set_VSlopeN):
                            OutType.append('C')
                            Sim_Time=Sim_Time+1
                            OutTim.append(Sim_Time)

                            if Sim_Distance<Section_SlopDistance[_i_Slope]:
                                OutSlope.append(Section_SlopList[_i_Slope])
                            else:
                                _i_Slope=_i_Slope+1
                                OutSlope.append(Section_SlopList[_i_Slope])
                        
                            OutSpeed.append(Sim_Section_Max_Speed)
                            Stop_Start_Sim_Distance=Stop_Start_Sim_Distance+Sim_Section_Max_Speed
                            Stop_Start_Sim_Time=Stop_Start_Sim_Time+1
                            Sim_Distance=Sim_Distance+Sim_Section_Max_Speed
                        Breaking_Time=0
                        #Stop start section deceleration part simulator 
                        while Stop_Start_Sim_Distance<Stop_Start_Distance:
                            OutType.append('C')
                            Sim_Time=Sim_Time+1
                            OutTim.append(Sim_Time)

                            if Sim_Distance<Section_SlopDistance[_i_Slope]:
                                OutSlope.append(Section_SlopList[_i_Slope])
                            else:
                                _i_Slope=_i_Slope+1
                                OutSlope.append(Section_SlopList[_i_Slope])

                            if Sim_Section_Max_Speed-Breaking_Time*Set_VSlopeN>0:
                                OutSpeed.append(Sim_Section_Max_Speed-Breaking_Time*Set_VSlopeN)
                                Stop_Start_Sim_Distance=Stop_Start_Sim_Distance+(Sim_Section_Max_Speed-Breaking_Time*Set_VSlopeN-0.5*Set_VSlopeN)
                                Sim_Distance=Sim_Distance+(Sim_Section_Max_Speed-Breaking_Time*Set_VSlopeN-0.5*Set_VSlopeN)
                            else:
                                OutSpeed.append(0)
                            Breaking_Time=Breaking_Time+1


                        Sim_Overall_Section_Distance=Sim_Overall_Section_Distance+Stop_Start_Sim_Distance
                        Collector_Working=0

                        #Collection section collector-is-working (say empty bins) part simulator
                        while Collector_Working<Collection_Work_Time:
                            OutType.append('P')
                            Sim_Time=Sim_Time+1
                            OutTim.append(Sim_Time)

                            OutSlope.append(OutSlope[len(OutSlope)-1])

                            OutSpeed.append(0)
                            Collector_Working=Collector_Working+1
                    Sim_Section_Distance_Left=Section_Distance-Sim_Overall_Section_Distance
                    #########################################################################UUUUUUUUPPPPPPPP
                    #########################################################################DDDDDDDDOOOOOOOOWWWWWNNNNN
                    #Calculate the last remaining section
                    if Sim_Section_Distance_Left>0:
                        #For the distance remaining calculate the speed if using a triangle approximation
                        Stop_Start_Speed=math.sqrt((2*Set_VSlopeP*Set_VSlopeN*Sim_Section_Distance_Left)/(Set_VSlopeP+Set_VSlopeN))
                        if Stop_Start_Speed < Max_Speed_Collection:
                            #1. cannot reach maximum speed
                            Sim_Section_Max_Speed=Stop_Start_Speed
                        else:
                            #2. maximum speed reached
                            Sim_Section_Max_Speed=Max_Speed_Collection
                        #Calculate the remaining section
                        Sim_Speed=0
                        Stop_Start_Sim_Distance=0
                        Stop_Start_Sim_Time=0
                        #Last stop start section acceleration part simulator
                        while Sim_Speed<=Sim_Section_Max_Speed:
                            OutType.append('C')
                            Sim_Time=Sim_Time+1
                            OutTim.append(Sim_Time)

                            if Sim_Distance<Section_SlopDistance[_i_Slope]:
                                OutSlope.append(Section_SlopList[_i_Slope])
                            else:
                                _i_Slope=_i_Slope+1
                                OutSlope.append(Section_SlopList[_i_Slope])

                            Sim_Speed=Stop_Start_Sim_Time*Set_VSlopeP
                            OutSpeed.append(Sim_Speed)
                            Stop_Start_Sim_Distance=Stop_Start_Sim_Distance+(Sim_Speed-0.5*Set_VSlopeP)

                            Sim_Distance=Sim_Distance+(Sim_Speed-0.5*Set_VSlopeP)

                            Stop_Start_Sim_Time=Stop_Start_Sim_Time+1
                        #Last stop start section const-speed part simulator  
                        while Stop_Start_Sim_Distance<Sim_Section_Distance_Left-(0.5*Sim_Section_Max_Speed*Sim_Section_Max_Speed/Set_VSlopeN):#Before the break distance
                            OutType.append('C')
                            Sim_Time=Sim_Time+1
                            OutTim.append(Sim_Time)

                            if Sim_Distance<Section_SlopDistance[_i_Slope]:
                                OutSlope.append(Section_SlopList[_i_Slope])
                            else:
                                _i_Slope=_i_Slope+1
                                OutSlope.append(Section_SlopList[_i_Slope])

                            OutSpeed.append(Sim_Section_Max_Speed)
                            Stop_Start_Sim_Distance=Stop_Start_Sim_Distance+Sim_Section_Max_Speed

                            Sim_Distance=Sim_Distance+Sim_Section_Max_Speed

                            Stop_Start_Sim_Time=Stop_Start_Sim_Time+1
                        Breaking_Time=0
                        Sim_Speed=0
                        #Last stop start section deceleration part simulator  
                        while Stop_Start_Sim_Distance<Sim_Section_Distance_Left:
                            OutType.append('C')
                            Sim_Time=Sim_Time+1
                            OutTim.append(Sim_Time)

                            if Sim_Distance<Section_SlopDistance[_i_Slope]:
                                OutSlope.append(Section_SlopList[_i_Slope])
                            else:
                                _i_Slope=_i_Slope+1
                                OutSlope.append(Section_SlopList[_i_Slope])

                            Sim_Speed=Sim_Section_Max_Speed-Breaking_Time*Set_VSlopeN
                            if Sim_Speed>0:
                                OutSpeed.append(Sim_Speed)
                                Stop_Start_Sim_Distance=Stop_Start_Sim_Distance+(Sim_Section_Max_Speed-Breaking_Time*Set_VSlopeN-0.5*Set_VSlopeN)

                                Sim_Distance=Sim_Distance+(Sim_Section_Max_Speed-Breaking_Time*Set_VSlopeN-0.5*Set_VSlopeN)

                            else:
                                OutSpeed.append(0)
                                Stop_Start_Sim_Distance=Sim_Section_Distance_Left#Quit when Speed < 0
                            Breaking_Time=Breaking_Time+1
                        #########################################################################UUUUUUUUPPPPPPPP
                Section_SlopDistance.clear()
                Section_SlopList.clear()
                Section_Distance=0
        #Debug plotter
        dataframe = pd.DataFrame({'OutType':OutType,'Time':OutTim,'OutSpeed':OutSpeed,'OutSlope':OutSlope})
        dataframe.to_csv(_FtoOpen+'_SimV_w_Slope.csv',index=True,sep=',')

def InsertDistance_CSV():
    files= os.listdir(CSV_path_Add_Distance) 
    for file in files: 
        if (os.path.isfile(CSV_path_Add_Distance+file)): 
            csv_File = pd.read_csv(CSV_path_Add_Distance+file, encoding='utf-8')
            csv_File.to_excel("Temp.xlsx", sheet_name='data')
            wb = xlrd.open_workbook("Temp.xlsx")
            _i=0
            print(CSV_path_Add_Distance+file.replace("csv","xls"))
            sheet1 = wb.sheet_by_index(0)
            _F_Lenth=sheet1.nrows
            workbook = xlwt.Workbook()
            sheet_new = workbook.add_sheet('data') 
            AccuDistance=0
            
            sheet_new.write(1, CSV_AccuDistanceatColumn, 0)
            sheet_new.write(1, CSV_DistanceatColumn, 0)
            while _i<=1:
                sheet_new.write(_i, 0, sheet1.cell(_i,0).value)
                sheet_new.write(_i, 1, sheet1.cell(_i,1).value)
                sheet_new.write(_i, 2, sheet1.cell(_i,2).value)
                sheet_new.write(_i, 3, sheet1.cell(_i,3).value)
                sheet_new.write(_i, 4, sheet1.cell(_i,4).value)
                sheet_new.write(_i, 5, sheet1.cell(_i,5).value)
                _i=_i+1
            sheet_new.write(0, CSV_AccuDistanceatColumn, "Distance(km)")
            sheet_new.write(0, CSV_DistanceatColumn, "Distance_interval (m)")
            sheet_new.write(0, CSV_DistanceatColumn+1, "Slope")
            sheet_new.write(1, CSV_DistanceatColumn+1, 0)

            _i=2
            while _i<_F_Lenth:
                sheet_new.write(_i, 0, sheet1.cell(_i,0).value)
                sheet_new.write(_i, 1, sheet1.cell(_i,1).value)
                sheet_new.write(_i, 2, sheet1.cell(_i,2).value)
                sheet_new.write(_i, 3, sheet1.cell(_i,3).value)
                sheet_new.write(_i, 4, sheet1.cell(_i,4).value)
                sheet_new.write(_i, 5, sheet1.cell(_i,5).value)
                coords_1 = (sheet1.cell(_i-1,CSV_LatitudeatColumn).value, sheet1.cell(_i-1,CSV_LongitudeatColumn).value)
                coords_2 = (sheet1.cell(_i,CSV_LatitudeatColumn).value, sheet1.cell(_i,CSV_LongitudeatColumn).value)
                Distance_i=geopy.distance.geodesic(coords_1, coords_2).m
                AccuDistance=AccuDistance+Distance_i
                sheet_new.write(_i, CSV_AccuDistanceatColumn, AccuDistance/1000)
                sheet_new.write(_i, CSV_DistanceatColumn, Distance_i)
                sheet_new.write(_i, CSV_DistanceatColumn+1, 0)
                _i=_i+1
            workbook.save(CSV_path_Add_Distance+"Output\\"+file.replace("csv","xls")) 

def InsertDistance():
    conn = pypyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\\Python\\RoadSimDom_Route.mdb;")
    #conn = adodbapi.connect("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Python\\RoadSimDom.mdb;")
    Crsr_Form = conn.cursor()
    Crsr_Contant = conn.cursor()
    Crsr_Dist = conn.cursor()  
    SQL = "SELECT name FROM MSYSOBJECTS WHERE ((TYPE=1) and flags=0 AND (name NOT LIKE '*MSys*'))"
    SQL_Rtn=Crsr_Form.execute(SQL).fetchall()
    for Forms in SQL_Rtn:
        print(str(Forms))
        if str(Forms) == "('Misc ',)" or str(Forms) == "(': ',)":
            continue
        Form_Name=str(Forms).replace("'","").replace("(","").replace(")","").replace(",","")
        SQL_Form=("select * from [" + Form_Name +"]")
        Form_Cont = Crsr_Contant.execute (SQL_Form).fetchall()
        try:
            SQL_Form=("alter table [" + Form_Name +"] add COLUMN Distance_m FLOAT")
            Crsr_Contant.execute (SQL_Form)
            SQL_Form=("ALTER TABLE [" + Form_Name +"] MODIFY COLUMN Distance_m decimal(10,5)")
            Crsr_Contant.execute (SQL_Form)
        except:
            Nothing=0
        edit_SQL="UPDATE [" + Form_Name + "] SET [Distance_m] = 0 WHERE [index] =  0"
        Crsr_Dist.execute(edit_SQL)
        Crsr_Dist.commit()
        _i=0
        _F_Lenth=len(Form_Cont)
        while _i<_F_Lenth:
            coords_1 = (Form_Cont[_i][DB_LatitudeatColumn], Form_Cont[_i][DB_LongitudeatColumn])
            coords_2 = (Form_Cont[_i+1][DB_LatitudeatColumn], Form_Cont[_i+1][DB_LongitudeatColumn])
            Distance_i=geopy.distance.geodesic(coords_1, coords_2).m
            edit_SQL="UPDATE [" + Form_Name + "] SET [Distance_m] = decimal(" + str(Distance_i) + ") WHERE [index] = '" + str(_i+1) + "'"
            Crsr_Dist.execute(edit_SQL)
            Crsr_Dist.commit()

def RouteBuilder():
    conn = pypyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\\Python\\RoadSimDom.mdb;")
    Crsr_Form = conn.cursor()
    Crsr_Contant = conn.cursor()
    Crsr_Elev = conn.cursor()  
    SQL = "SELECT name FROM MSYSOBJECTS WHERE ((TYPE=1) and flags=0 AND (name NOT LIKE '*MSys*'))"
    SQL_Rtn=Crsr_Form.execute(SQL).fetchall()

    connection_string = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Python\RoadSimDom_Route.mdb;'
    r'ExtendedAnsiSQL=1;'
    )
    connection_uri = f"access+pyodbc:///?odbc_connect={urllib.parse.quote_plus(connection_string)}"
    engine = create_engine(connection_uri)



    for Forms in SQL_Rtn:
        OutAddrAlt=[]
        OutAddrLat=[]
        OutAddrLong=[]
        OutAddrType=[]

        Addr_dict ={}
        Temp= {}
        _i=1

        print(str(Forms))
        if str(Forms).replace("(","").replace(")","").replace("'","").replace(",","").replace(" ","")== "Misc":
            continue
        Form_Name=str(Forms).replace("'","").replace("(","").replace(")","").replace(",","")
        SQL_Form=("select * from [" + Form_Name +"]")
        Form_Cont = Crsr_Contant.execute (SQL_Form).fetchall()


        _F_Lenth=len(Form_Cont)
        if _F_Lenth<Min_Form_Lenth:
            continue
        Address_Ignored=0
        while _i<_F_Lenth:

            Addr_Cell=Form_Cont[_i][DB_AddressatColumn]

            if Addr_Cell.find(',')<0: # deal with a address without "," like "65 Clarkehouse Road BROOMGROVE SHEFFIELD S102LG"
                print("No , -->" + Form_Cont[_i][DB_AddressatColumn])
                Address_Ignored=Address_Ignored+1
            
                Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                Crsr_Remove .execute(Remove_SQL)
                Crsr_Remove .commit()

                _i=_i+1
                continue
            if Addr_Cell.find('&')>0: # deal with a address with "&" , when a address with &, there are far too many possibilities so ignore it here
                print("Contains & -->" + Form_Cont[_i][DB_AddressatColumn])
                Address_Ignored=Address_Ignored+1

                Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                Crsr_Remove .execute(Remove_SQL)
                Crsr_Remove .commit()

                _i=_i+1
                continue
            Addr_Cell=Addr_Cell.upper()
            Addr_Cell=(Addr_Cell.replace(", ,", ""))
            Addr_Cell=(Addr_Cell.replace(":", ""))
            Addr_Cell=(Addr_Cell.replace("-", " TO "))
            Addr_Cell=(Addr_Cell.replace("  ", " "))
            Addr_Cell=(Addr_Cell.replace(")", ""))
            Addr_Cell=(Addr_Cell.replace("(", ""))
        
            if Addr_Cell.find(' TO ')>=0:     
                Try_Find_Number_1=Addr_Cell.split(',')[0].strip().split(' TO ')
                Try_Find_Number_2=Addr_Cell.split(',')[1].strip().split(' TO ')
                try:
                    House_Number_Min=int("".join(filter(str.isdigit, (Try_Find_Number_1[0].replace(" ", "").strip()))))
                    House_Number_Max=int("".join(filter(str.isdigit, (Try_Find_Number_1[1].replace(" ", "").strip()))))
                except:
                    try:
                        House_Number_Min=int("".join(filter(str.isdigit, (Try_Find_Number_2[0].replace(" ", "").strip()))))
                        House_Number_Max=int("".join(filter(str.isdigit, (Try_Find_Number_2[1].replace(" ", "").strip()))))
                    except:
                        print("Difficult (No Number) -->" + Form_Cont[_i][DB_AddressatColumn])
                        Address_Ignored=Address_Ignored+1

                        Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                        Crsr_Remove .execute(Remove_SQL)
                        Crsr_Remove .commit()

                        _i=_i+1
                        continue
                Cell_i=0
                Flag_Multi=0
                Addr_Cell_Split=Addr_Cell.split(',')
                while Cell_i<len(Addr_Cell_Split):  
                    if (Addr_Cell_Split[Cell_i].find('ROAD')+Addr_Cell_Split[Cell_i].find('AVENUE')+Addr_Cell_Split[Cell_i].find('LANE')+
                    Addr_Cell_Split[Cell_i].find('CRESCENT')+Addr_Cell_Split[Cell_i].find('SIDE')+Addr_Cell_Split[Cell_i].find('BANK')+
                    Addr_Cell_Split[Cell_i].find('STREET')+Addr_Cell_Split[Cell_i].find('CROFT')+Addr_Cell_Split[Cell_i].find('CLOSE')+
                    Addr_Cell_Split[Cell_i].find('DRIVE')+Addr_Cell_Split[Cell_i].find('VIEW')+Addr_Cell_Split[Cell_i].find('MEWS')+
                    Addr_Cell_Split[Cell_i].find('PLACE')+
                    Addr_Cell_Split[Cell_i].find('GROVE')+Addr_Cell_Split[Cell_i].find('TERRACE')+Addr_Cell_Split[Cell_i].find('WALK')+16>0):
                        break 
                    Cell_i=Cell_i+1
                if Cell_i>=len(Addr_Cell_Split):
                    print("Difficult (No Street) -->" + Form_Cont[_i][DB_AddressatColumn])
                    Address_Ignored=Address_Ignored+1

                    Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                    Crsr_Remove .execute(Remove_SQL)
                    Crsr_Remove .commit()

                    _i=_i+1
                    continue
                Temp_Sections=Addr_Cell_Split[Cell_i].strip().split(' ')
                if (Temp_Sections[0].isdigit()) or (len(char_regex.findall(Temp_Sections[0]))<=1 and len(num_regex.findall(Temp_Sections[0]))>=1) :
                    Temp_Num=int("".join(filter(str.isdigit, (Temp_Sections[0].replace(" ", "").strip()))))
                    Temp_i=1
                    Temp_Road=""
                    if Temp_Sections[1]=="TO":
                        House_Number_Min=Temp_Num
                        House_Number_Max=int("".join(filter(str.isdigit, (Temp_Sections[2].replace(" ", "").strip()))))
                        Temp_i=3
                        Flag_Multi=1
                    while Temp_i<len(Temp_Sections):
                        Temp_Road=Temp_Road + Temp_Sections[Temp_i] + " "
                        Temp_i=Temp_i+1
                    if Flag_Multi==1:
                        while House_Number_Min<=House_Number_Max:
                            addtodict3(Addr_dict,Temp_Road.replace("*", "").strip(),House_Number_Min,'RawAddress',Form_Cont[_i][DB_AddressatColumn])
                            addtodict3(Addr_dict,Temp_Road.replace("*", "").strip(),House_Number_Min,'Latitude',Form_Cont[_i][DB_LatitudeatColumn])
                            addtodict3(Addr_dict,Temp_Road.replace("*", "").strip(),House_Number_Min,'Longitude',Form_Cont[_i][DB_LongitudeatColumn])
                            addtodict3(Addr_dict,Temp_Road.replace("*", "").strip(),House_Number_Min,'Altitude', Form_Cont[_i][DB_AltitudeatColumn]) 
                            House_Number_Min=House_Number_Min+1
                    else:         
                        addtodict3(Addr_dict,Temp_Road.strip(),Temp_Num,'RawAddress',Form_Cont[_i][DB_AddressatColumn])
                        addtodict3(Addr_dict,Temp_Road.strip(),Temp_Num,'Latitude',Form_Cont[_i][DB_LatitudeatColumn])
                        addtodict3(Addr_dict,Temp_Road.strip(),Temp_Num,'Longitude',Form_Cont[_i][DB_LongitudeatColumn])
                        addtodict3(Addr_dict,Temp_Road.strip(),Temp_Num,'Altitude', Form_Cont[_i][DB_AltitudeatColumn]) 
                    _i=_i+1
                    continue  
                else:
                    if len(Addr_Cell_Split[Cell_i].split(' '))>4:
                        print("Difficult (>4) -->" + Form_Cont[_i][DB_AddressatColumn])
                        Address_Ignored=Address_Ignored+1

                        Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                        Crsr_Remove .execute(Remove_SQL)
                        Crsr_Remove .commit()


                        _i=_i+1
                        continue
                while House_Number_Min<=House_Number_Max:
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),House_Number_Min,'RawAddress',Form_Cont[_i][DB_AddressatColumn])
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),House_Number_Min,'Latitude',Form_Cont[_i][DB_LatitudeatColumn])
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),House_Number_Min,'Longitude',Form_Cont[_i][DB_LongitudeatColumn])
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),House_Number_Min,'Altitude', Form_Cont[_i][DB_AltitudeatColumn]) 
 
                    House_Number_Min=House_Number_Min+1
                _i=_i+1
            else:
                Address_Element=0
                Addr_O_Rd=Addr_Cell.split(',')[Address_Element].lstrip()
                if (Addr_O_Rd.find('FLAT')>=0) or (Addr_O_Rd.find('**')>=0):
                    if (Addr_O_Rd.strip().split(' ')[0].isdigit()) and (Addr_Cell.split(',')[1].lstrip().split(' ')[0].isdigit()==False):
                        Addr_Cell=Addr_O_Rd.replace("FLATS", "").strip().split(' ')[0] + " " + Addr_Cell.split(',')[1]+ ", " + Addr_Cell.split(',')[2]
                    else:
                        Address_Element=1
                        Addr_O_Rd=Addr_Cell.split(',')[Address_Element].strip()


                Addr_O_Rd_Next=Addr_Cell.split(',')[Address_Element+1].strip()

                try:
                    while Addr_O_Rd[0].isalpha() and Address_Element<(len(Addr_Cell.split(','))-2):
                        Address_Element=Address_Element+1
                        Addr_O_Rd=Addr_Cell.split(',')[Address_Element].lstrip()
                        Addr_O_Rd_Next=Addr_Cell.split(',')[Address_Element+1].lstrip()
                except:
                    print("Unknow -->" + Form_Cont[_i][DB_AddressatColumn])
                    Address_Ignored=Address_Ignored+1

                    Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                    Crsr_Remove .execute(Remove_SQL)
                    Crsr_Remove .commit()


                    _i=_i+1
                    continue

                if Addr_O_Rd.isdigit(): #Special step for something like 91, BRUNSWICK STREET
                    Addr_O_Rd=Addr_O_Rd + ' ' + Addr_Cell.split(',')[1]
                    Addr_O_Rd_Next=Addr_Cell.split(',')[2]       
       
                Addr_O_Rd=Addr_O_Rd.lstrip()
                Add_Num_1=(Addr_O_Rd.split(' ')[0])
                Add_Num=Add_Num_1.split('-')[0]
                Add_Num="".join(filter(str.isdigit, Add_Num))#remove any 12d or 34F's D and F
                if len(Add_Num)==0 or Add_Num[0].isalpha():#No Number then just record a single address at that road, this address will gave a number of 9999
                    Cell_i=0
                    Addr_Cell_Split=Addr_Cell.split(',')
                    while Cell_i<len(Addr_Cell_Split):  
                        if (Addr_Cell_Split[Cell_i].find('ROAD')+Addr_Cell_Split[Cell_i].find('AVENUE')+Addr_Cell_Split[Cell_i].find('LANE')+
                        Addr_Cell_Split[Cell_i].find('CRESCENT')+Addr_Cell_Split[Cell_i].find('SIDE')+Addr_Cell_Split[Cell_i].find('BANK')+
                        Addr_Cell_Split[Cell_i].find('STREET')+Addr_Cell_Split[Cell_i].find('CROFT')+Addr_Cell_Split[Cell_i].find('CLOSE')+
                        Addr_Cell_Split[Cell_i].find('DRIVE')+Addr_Cell_Split[Cell_i].find('VIEW')+Addr_Cell_Split[Cell_i].find('MEWS')+
                        Addr_Cell_Split[Cell_i].find('PLACE')+
                        Addr_Cell_Split[Cell_i].find('GROVE')+Addr_Cell_Split[Cell_i].find('TERRACE')+Addr_Cell_Split[Cell_i].find('WALK')+16>0):
                            break 
                        Cell_i=Cell_i+1
                    if Cell_i>=len(Addr_Cell_Split):
                        print("No Number (No Street Found) -->" +Form_Cont[_i][DB_AddressatColumn])
                        Address_Ignored=Address_Ignored+1

                        Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                        Crsr_Remove .execute(Remove_SQL)
                        Crsr_Remove .commit()

                        _i=_i+1
                        continue

                    if len(Addr_Cell_Split[Cell_i].split(' '))>4:
                        print("No Number (>4) -->" +Form_Cont[_i][DB_AddressatColumn])
                        Address_Ignored=Address_Ignored+1

                        Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                        Crsr_Remove .execute(Remove_SQL)
                        Crsr_Remove .commit()

                        _i=_i+1
                        continue
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),9999,'RawAddress',Form_Cont[_i][DB_AddressatColumn])
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),9999,'Latitude',Form_Cont[_i][DB_LatitudeatColumn])
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),9999,'Longitude',Form_Cont[_i][DB_LongitudeatColumn])
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),9999,'Altitude', Form_Cont[_i][DB_AltitudeatColumn]) 
                    _i=_i+1
                    continue
                else:
                    Add_Num=(int)(Add_Num)      
                Add_Pointer=1
                Add_Street=''
                while  len(Addr_O_Rd)-Add_Pointer>=0:
                    if Addr_O_Rd[len(Addr_O_Rd)-Add_Pointer].isdigit():
                        break
                    else:
                        Add_Street=Addr_O_Rd[len(Addr_O_Rd)-Add_Pointer] + ''+ Add_Street                 
                    Add_Pointer=Add_Pointer+1
                Add_Street=Add_Street.lstrip()
                if len(Add_Street)<7: # When Add_Street is empty, read the second element for Add_Street
                    Add_Street=Addr_O_Rd_Next.lstrip()
                if len(Add_Street)>7 and Add_Street[1] == ' ': #For something like: 82D Brunswick Street, this used to remove the 82D's D
                    Add_Street=Add_Street[2:len(Add_Street)]

                while len(Add_Street.split(' ')[0])==1:# remove "& E ABC Road" & E 
                    Add_Street=Add_Street[2:len(Add_Street)]# remove the first letter or symbol and a space

                if len(Add_Street)<7:
                    Address_Ignored=Address_Ignored+1

                    Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                    Crsr_Remove .execute(Remove_SQL)
                    Crsr_Remove .commit()

                    print(Form_Cont[_i][DB_AddressatColumn])
                    _i=_i+1
                    continue
                else:
                    addtodict3(Addr_dict,Add_Street.replace("*", "").strip(),Add_Num,'RawAddress',Form_Cont[_i][DB_AddressatColumn])
                    addtodict3(Addr_dict,Add_Street.replace("*", "").strip(),Add_Num,'Latitude',Form_Cont[_i][DB_LatitudeatColumn])
                    addtodict3(Addr_dict,Add_Street.replace("*", "").strip(),Add_Num,'Longitude',Form_Cont[_i][DB_LongitudeatColumn])
                    addtodict3(Addr_dict,Add_Street.replace("*", "").strip(),Add_Num,'Altitude', Form_Cont[_i][DB_AltitudeatColumn])    
                _i=_i+1
        print('Address_Ignored: ' + str(Address_Ignored))
        Num_Streets=len(list(Addr_dict))

        ##############################################################################################
        ##############################################################################################
        #Get the first street   
        ##############################################################################################  
        InitStreet=""
        Highest_Alt=0;

        for AStreet in Addr_dict:
            H_Num=list(Addr_dict[AStreet].keys()) #Get all house numbers
            if float(Addr_dict[AStreet][min(H_Num)]['Altitude'])>Highest_Alt:
                Highest_Alt=float(Addr_dict[AStreet][min(H_Num)]['Altitude'])
                InitStreet=AStreet
            if float(Addr_dict[AStreet][max(H_Num)]['Altitude'])>Highest_Alt:
                Highest_Alt=float(Addr_dict[AStreet][max(H_Num)]['Altitude'])
                InitStreet=AStreet
        print(InitStreet)
        print(Highest_Alt)


        a=list(Addr_dict[InitStreet].keys()) #Get all house numbers
        if Addr_dict[InitStreet][min(a)]['Altitude']>Addr_dict[InitStreet][max(a)]['Altitude']:
            a.sort(reverse=False)# Set starting point to Min(a) which is the number with highest Altitude+ '
            Next_st_Start_Lat=Addr_dict[InitStreet][min(a)]['Latitude']
            Next_st_Start_Long=Addr_dict[InitStreet][min(a)]['Longitude']
            Next_st_Stop_Lat=Addr_dict[InitStreet][max(a)]['Latitude']
            Next_st_Stop_Long=Addr_dict[InitStreet][max(a)]['Longitude']
        else:
            a.sort(reverse=True)# Set starting point to Max(a) which is the number with highest Altitude
            Next_st_Start_Lat=Addr_dict[InitStreet][max(a)]['Latitude']
            Next_st_Start_Long=Addr_dict[InitStreet][max(a)]['Longitude']
            Next_st_Stop_Lat=Addr_dict[InitStreet][min(a)]['Latitude']
            Next_st_Stop_Long=Addr_dict[InitStreet][min(a)]['Longitude']
        ##############################################################################################
        ##############################################################################################
        #Followings are from depot to the first street   
        ##############################################################################################   
        URL_Cruise = 'https://maps.googleapis.com/maps/api/directions/json?origin=' + str(Depot_Lat) + ',' + str(Depot_Long) + '&destination=' + str(Next_st_Start_Lat) + ',' + str(Next_st_Start_Long) + '&key=' + Google_API_Key   
        googleResponse = urllib.request.urlopen(URL_Cruise)
        jsonResponse = json.loads(googleResponse.read())
        googleResponse.close()
        Routes=jsonResponse['routes']
        Route_0=Routes[0]
        Leg_0=Route_0['legs']
        Steps=Leg_0[0]['steps']
        L=len(Steps)
        i=0
        while i < L:
            Cruise_Points=polyline.decode(Steps[i]['polyline']['points'])
            x=0
            while x < len(Cruise_Points):
                x_Lat=Cruise_Points[x][0]
                x_Long=Cruise_Points[x][1]
                OutAddrAlt.append('0')
                OutAddrLat.append(x_Lat)
                OutAddrLong.append(x_Long)
                OutAddrType.append('N')#Cruise from depot
                x=x+1
            i=i+1 
            OutAddrAlt.append('0')
            OutAddrLat.append(x_Lat)
            OutAddrLong.append(x_Long)
            OutAddrType.append('C')#Stop at each cross
        
        
        #######################################################################################
        #######################################################################################
        #Starting the 1st street  
        #######################################################################################  
        URL_Cruise = 'https://maps.googleapis.com/maps/api/directions/json?origin=' + str(Next_st_Start_Lat) + ',' + str(Next_st_Start_Long) + '&destination=' + str(Next_st_Stop_Lat) + ',' + str(Next_st_Stop_Long) + '&key=' + Google_API_Key    
        googleResponse = urllib.request.urlopen(URL_Cruise)
        jsonResponse = json.loads(googleResponse.read())
        googleResponse.close()
        Routes=jsonResponse['routes']
        Route_0=Routes[0]
        Leg_0=Route_0['legs']
        Steps=Leg_0[0]['steps']
        L=len(Steps)
        i=0
        while i < L:
            Cruise_Points=polyline.decode(Steps[i]['polyline']['points'])
            x=0
            while x < len(Cruise_Points):
                x_Lat=Cruise_Points[x][0]
                x_Long=Cruise_Points[x][1]
                OutAddrAlt.append('0')
                OutAddrLat.append(x_Lat)
                OutAddrLong.append(x_Long)
                OutAddrType.append('C')#Collection
                x=x+1
            i=i+1

        C_End_Lat=Next_st_Stop_Lat
        C_End_Long=Next_st_Stop_Long

        Addr_dict.pop(InitStreet)
        Street_done=1
        while Street_done<Num_Streets:
            Next_St=""
            Next_Max_Dist=9999;
            for All_Street in Addr_dict:
                    a=list(Addr_dict[All_Street].keys())
                    Min_St_Lat=Addr_dict[All_Street][min(a)]['Latitude']
                    Min_St_Long=Addr_dict[All_Street][min(a)]['Longitude']
                    Max_St_Lat=Addr_dict[All_Street][max(a)]['Latitude']
                    Max_St_Long=Addr_dict[All_Street][max(a)]['Longitude']
                    if Addr_dict[All_Street][min(a)]['Altitude']>Addr_dict[All_Street][max(a)]['Altitude']:

                        Next_Lat_Try=Min_St_Lat
                        Next_Long_Try=Min_St_Long
                    else:

                        Next_Lat_Try=Max_St_Lat
                        Next_Long_Try=Max_St_Long
                    URL_Distance = 'https://maps.googleapis.com/maps/api/directions/json?origin=' + str(C_End_Lat) + ',' + str(C_End_Long) + '&destination=' + str(Next_Lat_Try) + ',' + str(Next_Long_Try) + '&key=' + Google_API_Key  
                    googleResponse = urllib.request.urlopen(URL_Distance)
                    jsonResponse = json.loads(googleResponse.read())
                    googleResponse.close()
                    Routes=jsonResponse['routes']
                    Route_0=Routes[0]
                    Leg_0=Route_0['legs']
                    This_Dist=Leg_0[0]['distance']['value']/100 #convert to km

                    if This_Dist<Next_Max_Dist:
                        Next_St=All_Street
                        Next_Max_Dist=This_Dist
            print ("-------")

            print (Next_St)
            a=list(Addr_dict[Next_St].keys()) #Get all house numbers
            if Addr_dict[Next_St][min(a)]['Altitude']>Addr_dict[Next_St][max(a)]['Altitude']:
                a.sort(reverse=False)# Set starting point to Min(a) which is the number with highest Altitude
                Next_st_Start_Lat=Addr_dict[Next_St][min(a)]['Latitude']
                Next_st_Start_Long=Addr_dict[Next_St][min(a)]['Longitude']
                Next_st_Stop_Lat=Addr_dict[Next_St][max(a)]['Latitude']
                Next_st_Stop_Long=Addr_dict[Next_St][max(a)]['Longitude']
            else:
                a.sort(reverse=True)# Set starting point to Max(a) which is the number with highest Altitude
                Next_st_Start_Lat=Addr_dict[Next_St][max(a)]['Latitude']
                Next_st_Start_Long=Addr_dict[Next_St][max(a)]['Longitude']
                Next_st_Stop_Lat=Addr_dict[Next_St][min(a)]['Latitude']
                Next_st_Stop_Long=Addr_dict[Next_St][min(a)]['Longitude']
            #insert the cruise section
            URL_Cruise = 'https://maps.googleapis.com/maps/api/directions/json?origin=' + str(C_End_Lat) + ',' + str(C_End_Long) + '&destination=' + str(Next_st_Start_Lat) + ',' + str(Next_st_Start_Long) + '&key=' + Google_API_Key  
            googleResponse = urllib.request.urlopen(URL_Cruise)
            jsonResponse = json.loads(googleResponse.read())
            googleResponse.close()
            Routes=jsonResponse['routes']
            Route_0=Routes[0]
            Leg_0=Route_0['legs']
            Steps=Leg_0[0]['steps']
            L=len(Steps)
            i=0
            while i < L:
                Cruise_Points=polyline.decode(Steps[i]['polyline']['points'])
                x=0
                while x < len(Cruise_Points):
                    x_Lat=Cruise_Points[x][0]
                    x_Long=Cruise_Points[x][1]
                    OutAddrAlt.append('0')
                    OutAddrLat.append(x_Lat)
                    OutAddrLong.append(x_Long)
                    OutAddrType.append('N')#Cruise
                    x=x+1
                i=i+1
            if len(a)==1:
                for x in a:
                    OutAddrAlt.append(Addr_dict[Next_St][x]['Altitude'])
                    OutAddrLat.append(Addr_dict[Next_St][x]['Latitude'])
                    OutAddrLong.append(Addr_dict[Next_St][x]['Longitude'])
                    OutAddrType.append('C')#Collection
            else:
                URL_Cruise = 'https://maps.googleapis.com/maps/api/directions/json?origin=' + str(Next_st_Start_Lat) + ',' + str(Next_st_Start_Long) + '&destination=' + str(Next_st_Stop_Lat) + ',' + str(Next_st_Stop_Long) + '&key=' + Google_API_Key  
                googleResponse = urllib.request.urlopen(URL_Cruise)
                jsonResponse = json.loads(googleResponse.read())
                googleResponse.close()
                Routes=jsonResponse['routes']
                Route_0=Routes[0]
                Leg_0=Route_0['legs']
                Steps=Leg_0[0]['steps']
                L=len(Steps)
                i=0
                while i < L:
                    Cruise_Points=polyline.decode(Steps[i]['polyline']['points'])
                    x=0
                    while x < len(Cruise_Points):
                        x_Lat=Cruise_Points[x][0]
                        x_Long=Cruise_Points[x][1]
                        OutAddrAlt.append('0')
                        OutAddrLat.append(x_Lat)
                        OutAddrLong.append(x_Long)
                        OutAddrType.append('C')#Collection
                        x=x+1
                    i=i+1
            C_End_Lat=Next_st_Stop_Lat
            C_End_Long=Next_st_Stop_Long
            Addr_dict.pop(Next_St)
            Street_done=Street_done+1
    
        ##############################################################################################
        ##############################################################################################
        #Followings are from last street back to depot   
        ##############################################################################################   
        URL_Cruise = 'https://maps.googleapis.com/maps/api/directions/json?origin=' + str(C_End_Lat) + ',' + str(C_End_Long) + '&destination=' + str(Depot_Lat) + ',' + str(Depot_Long) + '&key=' + Google_API_Key  
        googleResponse = urllib.request.urlopen(URL_Cruise)
        jsonResponse = json.loads(googleResponse.read())
        googleResponse.close()
        Routes=jsonResponse['routes']
        Route_0=Routes[0]
        Leg_0=Route_0['legs']
        Steps=Leg_0[0]['steps']
        L=len(Steps)
        i=0
        while i < L:
            Cruise_Points=polyline.decode(Steps[i]['polyline']['points'])
            x=0
            while x < len(Cruise_Points):
                x_Lat=Cruise_Points[x][0]
                x_Long=Cruise_Points[x][1]
                OutAddrAlt.append('0')
                OutAddrLat.append(x_Lat)
                OutAddrLong.append(x_Long)
                OutAddrType.append('N')#Cruise to depot
                x=x+1
            i=i+1 
            OutAddrAlt.append('0')
            OutAddrLat.append(x_Lat)
            OutAddrLong.append(x_Long)
            OutAddrType.append('C')#Stop at each cross    
        dataframe = pd.DataFrame({'Type':OutAddrType,'Latitude':OutAddrLat,'Longitude':OutAddrLong, 'Altitude':OutAddrAlt})
        dataframe.to_sql(Form_Name, con=engine, if_exists='append')
        dataframe.to_csv(Form_Name.replace(":","-").replace("/","-")+ ".csv",index=True,sep=',')

def InsertElev():
    conn = pypyodbc.connect(r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\\Python\\RoadSimDom.mdb;")
    Crsr_Form = conn.cursor()
    Crsr_Contant = conn.cursor()
    Crsr_Elev = conn.cursor()  
    Crsr_Remove = conn.cursor()  
    SQL = "SELECT name FROM MSYSOBJECTS WHERE ((TYPE=1) and flags=0 AND (name NOT LIKE '*MSys*'))"
    SQL_Rtn=Crsr_Form.execute(SQL).fetchall()
    for Forms in SQL_Rtn:
        print(str(Forms))
        if str(Forms).replace("(","").replace(")","").replace("'","").replace(",","").replace(" ","")== "Misc":
            continue
        Form_Name=str(Forms).replace("'","").replace("(","").replace(")","").replace(",","")
        SQL_Form=("select * from [" + Form_Name +"]")
        Form_Cont = Crsr_Contant.execute (SQL_Form).fetchall()
        _i=0
        Addr_dict ={}
        _F_Lenth=len(Form_Cont)

        Address_Ignored=0
        while _i<_F_Lenth:

            Addr_Cell=Form_Cont[_i][DB_AddressatColumn]

            if Addr_Cell.find(',')<0: # deal with a address without "," like "65 Clarkehouse Road BROOMGROVE SHEFFIELD S102LG"
                print("No , -->" + Form_Cont[_i][DB_AddressatColumn])
                Address_Ignored=Address_Ignored+1
            
                Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                Crsr_Remove .execute(Remove_SQL)
                Crsr_Remove .commit()

                _i=_i+1
                continue
            if Addr_Cell.find('&')>0: # deal with a address with "&" , when a address with &, there are far too many possibilities so ignore it here
                print("Contains & -->" + Form_Cont[_i][DB_AddressatColumn])
                Address_Ignored=Address_Ignored+1

                Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                Crsr_Remove .execute(Remove_SQL)
                Crsr_Remove .commit()

                _i=_i+1
                continue
            Addr_Cell=Addr_Cell.upper()
            Addr_Cell=(Addr_Cell.replace(", ,", ""))
            Addr_Cell=(Addr_Cell.replace(":", ""))
            Addr_Cell=(Addr_Cell.replace("-", " TO "))
            Addr_Cell=(Addr_Cell.replace("  ", " "))
            Addr_Cell=(Addr_Cell.replace(")", ""))
            Addr_Cell=(Addr_Cell.replace("(", ""))
        
            if Addr_Cell.find(' TO ')>=0:     
                Try_Find_Number_1=Addr_Cell.split(',')[0].strip().split(' TO ')
                Try_Find_Number_2=Addr_Cell.split(',')[1].strip().split(' TO ')
                try:
                    House_Number_Min=int("".join(filter(str.isdigit, (Try_Find_Number_1[0].replace(" ", "").strip()))))
                    House_Number_Max=int("".join(filter(str.isdigit, (Try_Find_Number_1[1].replace(" ", "").strip()))))
                except:
                    try:
                        House_Number_Min=int("".join(filter(str.isdigit, (Try_Find_Number_2[0].replace(" ", "").strip()))))
                        House_Number_Max=int("".join(filter(str.isdigit, (Try_Find_Number_2[1].replace(" ", "").strip()))))
                    except:
                        print("Difficult (No Number) -->" + Form_Cont[_i][DB_AddressatColumn])
                        Address_Ignored=Address_Ignored+1

                        Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                        Crsr_Remove .execute(Remove_SQL)
                        Crsr_Remove .commit()

                        _i=_i+1
                        continue
                Cell_i=0
                Flag_Multi=0
                Addr_Cell_Split=Addr_Cell.split(',')
                while Cell_i<len(Addr_Cell_Split):  
                    if (Addr_Cell_Split[Cell_i].find('ROAD')+Addr_Cell_Split[Cell_i].find('AVENUE')+Addr_Cell_Split[Cell_i].find('LANE')+
                    Addr_Cell_Split[Cell_i].find('CRESCENT')+Addr_Cell_Split[Cell_i].find('SIDE')+Addr_Cell_Split[Cell_i].find('BANK')+
                    Addr_Cell_Split[Cell_i].find('STREET')+Addr_Cell_Split[Cell_i].find('CROFT')+Addr_Cell_Split[Cell_i].find('CLOSE')+
                    Addr_Cell_Split[Cell_i].find('DRIVE')+Addr_Cell_Split[Cell_i].find('VIEW')+Addr_Cell_Split[Cell_i].find('MEWS')+
                    Addr_Cell_Split[Cell_i].find('PLACE')+
                    Addr_Cell_Split[Cell_i].find('GROVE')+Addr_Cell_Split[Cell_i].find('TERRACE')+Addr_Cell_Split[Cell_i].find('WALK')+16>0):
                        break 
                    Cell_i=Cell_i+1
                if Cell_i>=len(Addr_Cell_Split):
                    print("Difficult (No Street) -->" + Form_Cont[_i][DB_AddressatColumn])
                    Address_Ignored=Address_Ignored+1

                    Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                    Crsr_Remove .execute(Remove_SQL)
                    Crsr_Remove .commit()

                    _i=_i+1
                    continue
                Temp_Sections=Addr_Cell_Split[Cell_i].strip().split(' ')
                if (Temp_Sections[0].isdigit()) or (len(char_regex.findall(Temp_Sections[0]))<=1 and len(num_regex.findall(Temp_Sections[0]))>=1) :
                    Temp_Num=int("".join(filter(str.isdigit, (Temp_Sections[0].replace(" ", "").strip()))))
                    Temp_i=1
                    Temp_Road=""
                    if Temp_Sections[1]=="TO":
                        House_Number_Min=Temp_Num
                        House_Number_Max=int("".join(filter(str.isdigit, (Temp_Sections[2].replace(" ", "").strip()))))
                        Temp_i=3
                        Flag_Multi=1
                    while Temp_i<len(Temp_Sections):
                        Temp_Road=Temp_Road + Temp_Sections[Temp_i] + " "
                        Temp_i=Temp_i+1
                    if Flag_Multi==1:
                        while House_Number_Min<=House_Number_Max:
                            addtodict3(Addr_dict,Temp_Road.replace("*", "").strip(),House_Number_Min,'RawAddress',Form_Cont[_i][DB_AddressatColumn])
                            addtodict3(Addr_dict,Temp_Road.replace("*", "").strip(),House_Number_Min,'Latitude',Form_Cont[_i][DB_LatitudeatColumn])
                            addtodict3(Addr_dict,Temp_Road.replace("*", "").strip(),House_Number_Min,'Longitude',Form_Cont[_i][DB_LongitudeatColumn])
                            addtodict3(Addr_dict,Temp_Road.replace("*", "").strip(),House_Number_Min,'Altitude',0) 
                            House_Number_Min=House_Number_Min+1
                    else:         
                        addtodict3(Addr_dict,Temp_Road.strip(),Temp_Num,'RawAddress',Form_Cont[_i][DB_AddressatColumn])
                        addtodict3(Addr_dict,Temp_Road.strip(),Temp_Num,'Latitude',Form_Cont[_i][DB_LatitudeatColumn])
                        addtodict3(Addr_dict,Temp_Road.strip(),Temp_Num,'Longitude',Form_Cont[_i][DB_LongitudeatColumn])
                        addtodict3(Addr_dict,Temp_Road.strip(),Temp_Num,'Altitude',0) 
                    _i=_i+1
                    continue  
                else:
                    if len(Addr_Cell_Split[Cell_i].split(' '))>4:
                        print("Difficult (>4) -->" + Form_Cont[_i][DB_AddressatColumn])
                        Address_Ignored=Address_Ignored+1

                        Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                        Crsr_Remove .execute(Remove_SQL)
                        Crsr_Remove .commit()


                        _i=_i+1
                        continue
                while House_Number_Min<=House_Number_Max:
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),House_Number_Min,'RawAddress',Form_Cont[_i][DB_AddressatColumn])
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),House_Number_Min,'Latitude',Form_Cont[_i][DB_LatitudeatColumn])
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),House_Number_Min,'Longitude',Form_Cont[_i][DB_LongitudeatColumn])
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),House_Number_Min,'Altitude',0) 
 
                    House_Number_Min=House_Number_Min+1
                _i=_i+1
            else:
                Address_Element=0
                Addr_O_Rd=Addr_Cell.split(',')[Address_Element].lstrip()
                if (Addr_O_Rd.find('FLAT')>=0) or (Addr_O_Rd.find('**')>=0):
                    if (Addr_O_Rd.strip().split(' ')[0].isdigit()) and (Addr_Cell.split(',')[1].lstrip().split(' ')[0].isdigit()==False):
                        Addr_Cell=Addr_O_Rd.replace("FLATS", "").strip().split(' ')[0] + " " + Addr_Cell.split(',')[1]+ ", " + Addr_Cell.split(',')[2]
                    else:
                        Address_Element=1
                        Addr_O_Rd=Addr_Cell.split(',')[Address_Element].strip()


                Addr_O_Rd_Next=Addr_Cell.split(',')[Address_Element+1].strip()

                try:
                    while Addr_O_Rd[0].isalpha() and Address_Element<(len(Addr_Cell.split(','))-2):
                        Address_Element=Address_Element+1
                        Addr_O_Rd=Addr_Cell.split(',')[Address_Element].lstrip()
                        Addr_O_Rd_Next=Addr_Cell.split(',')[Address_Element+1].lstrip()
                except:
                    print("Unknow -->" + Form_Cont[_i][DB_AddressatColumn])
                    Address_Ignored=Address_Ignored+1

                    Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                    Crsr_Remove .execute(Remove_SQL)
                    Crsr_Remove .commit()


                    _i=_i+1
                    continue

                if Addr_O_Rd.isdigit(): #Special step for something like 91, BRUNSWICK STREET
                    Addr_O_Rd=Addr_O_Rd + ' ' + Addr_Cell.split(',')[1]
                    Addr_O_Rd_Next=Addr_Cell.split(',')[2]       
       
                Addr_O_Rd=Addr_O_Rd.lstrip()
                Add_Num_1=(Addr_O_Rd.split(' ')[0])
                Add_Num=Add_Num_1.split('-')[0]
                #Add_Num=Add_Num.replace("A", "").replace("B", "").replace("C", "").replace("D", "").replace("E", "").replace("F", "").replace("G", "").replace("H", "")
                Add_Num="".join(filter(str.isdigit, Add_Num))#remove any 12d or 34F's D and F
                if len(Add_Num)==0 or Add_Num[0].isalpha():#No Number then just record a single address at that road, this address will gave a number of 9999
                    #print("No Number -->" +Form_Cont[_i][DB_AddressatColumn])
                    Cell_i=0
                    Addr_Cell_Split=Addr_Cell.split(',')
                    while Cell_i<len(Addr_Cell_Split):  
                        if (Addr_Cell_Split[Cell_i].find('ROAD')+Addr_Cell_Split[Cell_i].find('AVENUE')+Addr_Cell_Split[Cell_i].find('LANE')+
                        Addr_Cell_Split[Cell_i].find('CRESCENT')+Addr_Cell_Split[Cell_i].find('SIDE')+Addr_Cell_Split[Cell_i].find('BANK')+
                        Addr_Cell_Split[Cell_i].find('STREET')+Addr_Cell_Split[Cell_i].find('CROFT')+Addr_Cell_Split[Cell_i].find('CLOSE')+
                        Addr_Cell_Split[Cell_i].find('DRIVE')+Addr_Cell_Split[Cell_i].find('VIEW')+Addr_Cell_Split[Cell_i].find('MEWS')+
                        Addr_Cell_Split[Cell_i].find('PLACE')+
                        Addr_Cell_Split[Cell_i].find('GROVE')+Addr_Cell_Split[Cell_i].find('TERRACE')+Addr_Cell_Split[Cell_i].find('WALK')+16>0):
                            break 
                        Cell_i=Cell_i+1
                    if Cell_i>=len(Addr_Cell_Split):
                        print("No Number (No Street Found) -->" +Form_Cont[_i][DB_AddressatColumn])
                        Address_Ignored=Address_Ignored+1

                        Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                        Crsr_Remove .execute(Remove_SQL)
                        Crsr_Remove .commit()

                        _i=_i+1
                        continue

                    if len(Addr_Cell_Split[Cell_i].split(' '))>4:
                        print("No Number (>4) -->" +Form_Cont[_i][DB_AddressatColumn])
                        Address_Ignored=Address_Ignored+1

                        Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                        Crsr_Remove .execute(Remove_SQL)
                        Crsr_Remove .commit()

                        _i=_i+1
                        continue
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),9999,'RawAddress',Form_Cont[_i][DB_AddressatColumn])
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),9999,'Latitude',Form_Cont[_i][DB_LatitudeatColumn])
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),9999,'Longitude',Form_Cont[_i][DB_LongitudeatColumn])
                    addtodict3(Addr_dict,Addr_Cell_Split[Cell_i].replace("*", "").strip(),9999,'Altitude',0) 
                    _i=_i+1
                    continue
                else:
                    Add_Num=(int)(Add_Num)      
                Add_Pointer=1
                Add_Street=''
                while  len(Addr_O_Rd)-Add_Pointer>=0:
                    if Addr_O_Rd[len(Addr_O_Rd)-Add_Pointer].isdigit():
                        break
                    else:
                        Add_Street=Addr_O_Rd[len(Addr_O_Rd)-Add_Pointer] + ''+ Add_Street                 
                    Add_Pointer=Add_Pointer+1
                Add_Street=Add_Street.lstrip()
                if len(Add_Street)<7: # When Add_Street is empty, read the second element for Add_Street
                    Add_Street=Addr_O_Rd_Next.lstrip()
                if len(Add_Street)>7 and Add_Street[1] == ' ': #When something like: 82D Brunswick Street, this used to remove the 82D's D
                    Add_Street=Add_Street[2:len(Add_Street)]

                while len(Add_Street.split(' ')[0])==1:# remove "& E ABC Road" & E 
                    Add_Street=Add_Street[2:len(Add_Street)]# remove the first letter or symbol and a space

                if len(Add_Street)<7:
                    Address_Ignored=Address_Ignored+1

                    Remove_SQL="delete FROM[" + Form_Name + "] WHERE [Address_ID] = " + str(_i+1) + ""
                    Crsr_Remove .execute(Remove_SQL)
                    Crsr_Remove .commit()

                    print(Form_Cont[_i][DB_AddressatColumn])
                    _i=_i+1
                    continue
                else:
                    addtodict3(Addr_dict,Add_Street.replace("*", "").strip(),Add_Num,'RawAddress',Form_Cont[_i][DB_AddressatColumn])
                    addtodict3(Addr_dict,Add_Street.replace("*", "").strip(),Add_Num,'Latitude',Form_Cont[_i][DB_LatitudeatColumn])
                    addtodict3(Addr_dict,Add_Street.replace("*", "").strip(),Add_Num,'Longitude',Form_Cont[_i][DB_LongitudeatColumn])
                    addtodict3(Addr_dict,Add_Street.replace("*", "").strip(),Add_Num,'Altitude',0)    
                _i=_i+1
        print('Address_Ignored: ' + str(Address_Ignored))
        Num_Streets=len(list(Addr_dict))
        ##############################################################################################
        ##############################################################################################
        #Insert Elevation data for the Min and Max number of each street    
        ##############################################################################################     
        for All_Street in Addr_dict:
            a=list(Addr_dict[All_Street].keys()) #Get all house numbers
            googleResponse=urllib.request.urlopen('https://maps.googleapis.com/maps/api/elevation/json?locations=' + str(Addr_dict[All_Street][min(a)]['Latitude']) + ',' + str(Addr_dict[All_Street][min(a)]['Longitude']) + '&key=AIzaSyDbMAsjgMjp-fdgBmapvlYFYfYutAPJDUM') 
            Elevation_Min = json.loads(googleResponse.read())['results'][0]['elevation']
            googleResponse.close()
            Addr_dict[All_Street][min(a)]['Altitude']=Elevation_Min
            edit_SQL="UPDATE [" + Form_Name + "] SET [Elevation] = " + str(Elevation_Min) + " WHERE [Address] = '" + str(Addr_dict[All_Street][min(a)]['RawAddress'].replace("'","''")) + "'"
            Crsr_Elev.execute(edit_SQL)
            googleResponse=urllib.request.urlopen('https://maps.googleapis.com/maps/api/elevation/json?locations=' + str(Addr_dict[All_Street][max(a)]['Latitude']) + ',' + str(Addr_dict[All_Street][max(a)]['Longitude']) + '&key=AIzaSyDbMAsjgMjp-fdgBmapvlYFYfYutAPJDUM') 
            Elevation_Max = json.loads(googleResponse.read())['results'][0]['elevation']
            googleResponse.close()
            Addr_dict[All_Street][max(a)]['Altitude']=Elevation_Max
            edit_SQL="UPDATE [" + Form_Name + "] SET [Elevation] = " + str(Elevation_Max) + " WHERE [Address] = '" + str(Addr_dict[All_Street][max(a)]['RawAddress'].replace("'","''")) + "'"
            Crsr_Elev.execute(edit_SQL)
            Crsr_Elev.commit()


def DatabaseBuilder():
    wb = xlrd.open_workbook(filename=Bin_file)
    print(wb.sheet_names())
    sheet1 = wb.sheet_by_index(0)
    _F_Lenth=sheet1.nrows

    conn_str = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Python\\RoadSimDom.mdb;")
    conn = adodbapi.connect(conn_str)
    crsr = conn.cursor()
    SQL1 = "CREATE TABLE [Misc] ([ID] AUTOINCREMENT PRIMARY KEY , [Name] VARCHAR(20),[Val] VARCHAR(50))"
    crsr.execute(SQL1)
    SQL2 = "INSERT INTO [Misc] ([Name],[Val]) VALUES ('Depot_Lat','" + str(Depot_Lat) + "')"
    crsr.execute(SQL2)
    SQL2 = "INSERT INTO [Misc] ([Name],[Val]) VALUES ('Depot_Long','" + str(Depot_Long) + "')"
    crsr.execute(SQL2)
    conn.commit()


    _i=1
    while _i<_F_Lenth:
        DomGroupName_i=sheet1.cell(_i,DomGroupAtColumn).value
        if DomGroupName_i.value.strip()==":":
            _i=_i+1
            continue
        DomAddr_i=str(sheet1.cell(_i,DomAddrAtColumn).value).replace("'","''")
        Postcode_i=str(sheet1.cell(_i,PostcodeAtColumn).value)
        Latitude_i=str(Form_Cont[_i][DB_LatitudeatColumn])
        Longitude_i=str(Form_Cont[_i][DB_LongitudeatColumn])
        if Postcode_i == "":
            Postcode_i=0.00  
        if Latitude_i == "NULL":# or type(Latitude_i) != float:
            Postcode_i=0.00  
        if Longitude_i == "NULL":#or type(Longitude_i) != float:
            Postcode_i=0.00  
        Elevation_i=str(0) 
        if DomGroupName_i in DomGroupName:
            SQL_Insert =    "INSERT INTO [" + DomGroupName_i + "] ([Address],[Postcode],[Latitude],[Longitude],[Elevation]) VALUES ('" + \
                             DomAddr_i + " ','" + \
                             Postcode_i + " ','" + \
                             Latitude_i + " ','" + \
                             Longitude_i + " ','" + \
                             Elevation_i + "')"
            crsr.execute(SQL_Insert)
        else:
            DomGroupName.append(DomGroupName_i)
            SQL_Creat = "CREATE TABLE [" + DomGroupName_i + "] ([Address_ID] AUTOINCREMENT PRIMARY KEY , [Address] VARCHAR(255),[Postcode] VARCHAR(20), [Latitude] DECIMAL(12,8), [Longitude] DECIMAL(12,8), [Elevation] DECIMAL(12,8))"
            crsr.execute(SQL_Creat)
            SQL_Insert =    "INSERT INTO [" + DomGroupName_i + "] ([Address],[Postcode],[Latitude],[Longitude],[Elevation]) VALUES ('" + \
                             DomAddr_i + " ','" + \
                             Postcode_i + " ','" + \
                             Latitude_i + " ','" + \
                             Longitude_i + " ','" + \
                             Elevation_i + "')"
            crsr.execute(SQL_Insert)
        _i=_i+1
        conn.commit()
        print(_i)
    crsr.close()
    conn.close()
    
    
    END=1
     




#pypyodbc.win_create_mdb('C:\\Python\\RoadSimDom.mdb')  
 
#DatabaseBuilder()
#InsertElev()
#pypyodbc.win_create_mdb('C:\\Python\\RoadSimDom_Route.mdb') 
#RouteBuilder()
#InsertDistance()
InsertDistance_CSV()
#SpeedProfBuilder()

