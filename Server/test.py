import pyodbc
#### GLOBAL INFO> ABOUT SERVER ####
fs_ip = open("server.txt", 'r')
sqlip_line = fs_ip.readline()
sqlport_line = fs_ip.readline()
sqluser_line = fs_ip.readline()
sqlpass_line = fs_ip.readline()
sqldb_line = fs_ip.readline()

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

print("---SQL_Server---")
print("192.168.1.5")
print(sqlPort)
print(sqlUsername)
print(sqlPassword)
print(sqlDatabase)
print("##############")

connSQL = 'DRIVER={SQL Server};SERVER='+sqlServer+';PORT='+sqlPort+';DATABASE='+ sqlDatabase+';UID='+sqlUsername+';PWD='+sqlPassword
connnection = pyodbc.connect('DRIVER={SQL Server};SERVER='+sqlServer+';PORT='+sqlPort+';DATABASE='+ sqlDatabase+';UID='+sqlUsername+';PWD='+sqlPassword)
print("ok")
cur = connnection.cursor()
checkTable = cur.tables(table='votes').fetchone()
print(checkTable)