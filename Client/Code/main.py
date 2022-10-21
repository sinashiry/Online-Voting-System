0# -*- coding: UTF-8 -*-

#####################################################
#    Project:     Online Voting System v2           #
#    Designer:    Hosein AlamShahi                  #
#    Programmer:  Sina Shiry                        #
#    Date:        2019 Mar 03                       #
#    For:         Refinery Of Tabriz                #
#####################################################


# Config and import needed Libraries
###
import RPi.GPIO as GPIO
###
from picamera import PiCamera
from PIL import Image
###
import time
import threading
import jdatetime
###
import sqlite3
import pyodbc
###
from os.path import exists
import os
import subprocess

months = {"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"May","06":"Jun",
          "07":"Jul","08":"Aug","09":"Sep","10":"Oct","11":"Nov","12":"Dec"}

#wait for booting up
time.sleep(3)
# Configure time location
jdatetime.set_locale('fa_IR')

# Configure GPIO Pin(s)
P1_Button = 17
P2_Button = 22
P3_Button = 23
P4_Button = 27
P5_Button = 24
# Setup GPIOs For get output and input from it
GPIO.setmode(GPIO.BCM)
GPIO.setup(P1_Button, GPIO.IN, pull_up_down = GPIO.PUD_UP)
GPIO.setup(P2_Button, GPIO.IN, pull_up_down = GPIO.PUD_UP)
GPIO.setup(P3_Button, GPIO.IN, pull_up_down = GPIO.PUD_UP)
GPIO.setup(P4_Button, GPIO.IN, pull_up_down = GPIO.PUD_UP)
GPIO.setup(P5_Button, GPIO.IN, pull_up_down = GPIO.PUD_UP)

# Global Variable(s)

# Get Static ip from GUI settings
f_ip = open("/etc/dhcpcd.conf", 'r').read()
ip_addr = f_ip[int(f_ip.find("interface eth0\ninform"))+22:int(f_ip.find("interface eth0\ninform"))+37]
poss = ip_addr.find("\n")
if poss == -1:
    ip_addr = ip_addr.replace("\n","")
    ip_addr = ip_addr.replace(" ","")
else:
    ip_addr = ip_addr[0:poss]
    ip_addr = ip_addr.replace("\n","")
    ip_addr = ip_addr.replace(" ","")

#Information about SQL SERVER
fs_ip = open("/home/pi/Desktop/server.txt", 'r')
device_line = fs_ip.readline()
sqlip_line = fs_ip.readline()
sqlport_line = fs_ip.readline()
sqluser_line = fs_ip.readline()
sqlpass_line = fs_ip.readline()
sqldb_line = fs_ip.readline()

deviceName = device_line[3:]
deviceName = deviceName.replace("\n","")
deviceName = deviceName.replace(" ","")
deviceName = "Device_" + deviceName

sqlServer = sqlip_line[3:]
sqlServer = sqlServer.replace("\n","")
sqlServer = sqlServer.replace(" ","")

sqlPort = sqlport_line[3:]
sqlPort = sqlPort.replace("\n","")
sqlPort = sqlPort.replace(" ","")

sqlUsername = sqluser_line[3:]
sqlUsername = sqlUsername.replace("\n","")
sqlUsername = sqlUsername.replace(" ","")

sqlPassword = sqlpass_line[3:]
sqlPassword = sqlPassword.replace("\n","")
sqlPassword = sqlPassword.replace(" ","")

sqlDatabase = sqldb_line[3:]
sqlDatabase = sqlDatabase.replace("\n","")
sqlDatabase = sqlDatabase.replace(" ","")

print("---DeviceName---")
print(deviceName)
print("---SQL_Server---")
print(sqlServer)
print(sqlPort)
print(sqlUsername)
print(sqlPassword)
print(sqlDatabase)
print("##############")
###

# Thread Duration for sync Time of RPi
set_time_thread = 10
set_time_thread_per = 3600

def set_datetime():
    try:
        output = pyodbc.connect('DRIVER={FreeTDS};SERVER='+sqlServer+';PORT='+sqlPort+';DATABASE='+
                                sqlDatabase+';UID='+sqlUsername+';PWD='+sqlPassword)
        cur = output.cursor()
        cur.execute("SELECT SYSDATETIME()").description
        serverDatetime = cur.fetchall()[0][0]
        now = serverDatetime[8:10] + " " + months[serverDatetime[5:7]] + " "
        now = now + serverDatetime[0:4] + " " + serverDatetime[11:19]
        output = subprocess.Popen(["sudo","date","-s",now],stdout = subprocess.PIPE).communicate()[0]
        print("#########################")
        print ("RPi Time same as SQL NTP Server")
        cur.close
        threading.Timer(set_time_thread_per, set_datetime).start()
    except:
        print("#########################")
        print ("RPi Cant Get SQL NTP Server")
        threading.Timer(set_time_thread, set_datetime).start()

def find_now_poll():
    try:
        output = pyodbc.connect('DRIVER={FreeTDS};SERVER='+sqlServer+';PORT='+sqlPort+';DATABASE='+
                                sqlDatabase+';UID='+sqlUsername+';PWD='+sqlPassword)
        cur = output.cursor()
        cur.execute("SELECT PATH_NUM FROM POLL_QUESTION WHERE ACT_VOTE = '*'")
        avt_vote = cur.fetchall()
        if len(avt_vote) != 0:
            return str(avt_vote[0][0])    
        else:
            return False
    except:
        print("#########################")
        print("Cant Connect to SQL Server")


def create_name():
    file_path = "/home/pi/Desktop/Database/Pictures/"
    date_ = str(jdatetime.datetime.now())[:10]
    time_ = str(jdatetime.datetime.now())[11:-7]
    time_win = time_.replace(":", "_")
    file_path = file_path + date_.replace("-", "") + "_" + time_.replace(":", "") + ".jpg"
    return file_path, date_, time_,

def take_picture(button_ , datetime_, filepath):
    camera = PiCamera()
    #print("1")
    camera.resolution = (640, 480)
    #print("2")
    camera.annotate_text = button_ + " (" + datetime_ + ")"
    #print("3")
    camera.capture(filepath)
    #print("4")
    camera.close()
    #print("5")
    print("Pic.: Taked")
    # Reduce image size
    cache = Image.open(filepath)
    #cache = cache.resize(480,640)
    cache.save(filepath,quality=85)
    print("Pic. Resized")
    

def record_vote(button):
    # Check New POLL
    #print("1-1")
    # Get time, date and file path 
    file_path, date_, time_ = create_name()
    #print("1-2")
    # Take a picture
    take_picture(button, date_ + "  " + time_, file_path)
    #print("1-3")
    ##############################
    # Record vote on SQL Server
    try:
        output = pyodbc.connect('DRIVER={FreeTDS};SERVER='+sqlServer+';PORT='+sqlPort+';DATABASE='+
                                sqlDatabase+';UID='+sqlUsername+';PWD='+sqlPassword)
        cur = output.cursor()
        poll_now = find_now_poll()
        if poll_now == False:
            raise IOError
        
        # Get which Voting in Active
        
        command = "INSERT INTO vote_" + poll_now + " (VOTE, DATE, TIME, DIVC, IMAG) "
        command = command + "VALUES ('"+button+"','"+date_+"','"+time_+"','" + deviceName + "'," + "?" + ")"
        #Get picture for sending
        data = open(file_path,'rb').read()
        
        cur.execute(command,(pyodbc.Binary(data),))
        cur.commit()
        cur.close
        
        print("#########################")
        print("Vote stored in SQL server")
        ##############################
        print("Vote Recorded")
    except:
        print("#########################")
        print("Can't connect SQL server")


try:
	print("program Running Successfully")
	set_datetime()
	while (True):
		a = GPIO.input(P1_Button)
		b = GPIO.input(P2_Button)
		c = GPIO.input(P3_Button)
		d = GPIO.input(P4_Button)
		e = GPIO.input(P5_Button)
		if a == False:
			a = True
			print("########################")
			print("Vote: Number_01")
			print("D&T : " + str(jdatetime.datetime.now())[:-7])
			record_vote("BUTTON_1")
			time.sleep(3)
		if b == False:
			b = True
			print("########################")
			print("Vote: Number_02")
			print("D&T : " + str(jdatetime.datetime.now())[:-7])
			record_vote("BUTTON_2")
			time.sleep(3)
		if c == False:
			c = True
			print("########################")
			print("Vote: Number_03")
			print("D&T : " + str(jdatetime.datetime.now())[:-7])
			record_vote("BUTTON_3")
			time.sleep(3)
		if d == False:
			d = True
			print("########################")
			print("Vote: Number_04")
			print("D&T : " + str(jdatetime.datetime.now())[:-7])
			record_vote("BUTTON_4")
			time.sleep(3)
		if e == False:
			e = True
			print("########################")
			print("Vote: Number_05")
			print("D&T : " + str(jdatetime.datetime.now())[:-7])
			record_vote("BUTTON_5")
			time.sleep(3)
except:
    print("error")
    output = subprocess.Popen(["sudo","reboot"],stdout = subprocess.PIPE).communicate()[0]

finally:
    #camera.close()
    GPIO.cleanup()
    print("Cleaning Up!")
