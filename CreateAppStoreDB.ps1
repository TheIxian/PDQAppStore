# Set the preferred SQLite executable location
$sqLitePath = 'sqlite3.exe'

# AppStore DB Path
$appStoreDbPath = 'as.db'

# formatted date string is inserted in to the database
$formattedDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

## as.db schema
#
# Table: Packages
# Column1 = [INTEGER PRIMARY KEY] PackageID
# Column2 = [TEXT] PackageName
# Column3 = [TEXT] PdqPackageName
# Column4 = [INTEGER] CategoryID
# Column5 = [INTEGER] PackageEnabled
# Column6 = [TEXT] PackageDescription
# Column7 = [INTEGER] InteractiveOnly
#
# Table: Categories
# Column1 = [INTEGER PRIMARY KEY] CategoryID
# Column2 = [TEXT] CategoryName
#
# Table: Deployments
# Column1 = [INTEGER PRIMARY KEY] QueueID
# Column2 = [TEXT] ComputerName
# Column3 = [TEXT] UserName
# Column4 = [INTEGER] PackageID
# Column5 = [INTEGER] PdqDeploymentID
# Column6 = [INTEGER] IsScheduled
# Column7 = [TEXT] ScheduledTime
#
# Table: Log
# Column1 = [INTEGER PRIMARY KEY] LogID
# Column2 = [TEXT] ComputerName
# Column3 = [TEXT] TimeStamp
# Column4 = [TEXT] Message

& $sqLitePath $appStoreDbPath "CREATE TABLE IF NOT EXISTS Packages 
									(PackageID INTEGER PRIMARY KEY,
									 PackageName TEXT,
									 PdqPackageName TEXT,
									 CategoryID INTEGER,
									 PackageEnabled INTEGER,
									 PackageDescription TEXT,
									 InteractiveOnly INTEGER)"

& $sqLitePath $appStoreDbPath "INSERT INTO Packages VALUES (0,'All',null,0,0,null,0)"
									 
& $sqLitePath $appStoreDbPath "CREATE TABLE IF NOT EXISTS Categories 
									(CategoryID INTEGER PRIMARY KEY,
									 CategoryName TEXT)"

& $sqLitePath $appStoreDbPath "INSERT INTO Categories VALUES (0,'All')"
									 
& $sqLitePath $appStoreDbPath "CREATE TABLE IF NOT EXISTS Deployments 
									(QueueID INTEGER PRIMARY KEY,
									 ComputerName TEXT,
									 UserName TEXT,
									 PackageID INTEGER,
									 PdqDeploymentID INTEGER,
									 IsScheduled INTEGER,
									 ScheduledTime TEXT)"
									 
& $sqLitePath $appStoreDbPath "INSERT INTO Deployments VALUES (0,'INITIALIZE','INITIALIZE',0,null,0,null)"
									 
& $sqLitePath $appStoreDbPath "CREATE TABLE IF NOT EXISTS Log
									(LogID INTEGER PRIMARY KEY, 
									 ComputerName TEXT,
									 TimeStamp TEXT,
									 Message TEXT)"
									 
& $sqLitePath $appStoreDbPath "INSERT INTO Log VALUES (0,'INITIALIZE','$formattedDateTime','Initial import of all PDQ packages')"
