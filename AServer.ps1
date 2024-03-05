###################################################################################
#
# PDQ AppStore Server script
#
###################################################################################
# CHANGE LOG
###################################################################################
# 
#
###################################################################################
###################################################################################
# VARIABLES
#

# Local SQLite executable location
$sqLitePath = 'sqlite3.exe'

# AppStore DB path
$appStoreDbPath = 'as.db'

# PDQ Deploy executable path
$pdqDeployExePath = 'C:\Program Files (x86)\Admin Arsenal\PDQ Deploy\PDQDeploy.exe'

# Write log to console
$logToConsole = $False

# Send e-mail notifications
$sendEmailNotifications = $True

# SMTP Server
$smtpServer = '127.0.0.1'

# SMTP use authentication?
$smtpUseAuth = $False

# SMTP User Name
$smtpUsername = 'noreply@domain.com'

# SMTP Password
$smtpPassword = $Null

# Admin e-mail address
$smtpAdminAddress = 'alert@domain.com'

# Mail From address
$smtpMailFrom = '"PDQ AppStore" <' + $smtpUsername + '>'

#
# END VARIABLES
###################################################################################
###################################################################################
# INITIALIZE
#

# Set up SMTP credential object
if ($smtpUseAuth) {
	$smtpPWord = ConvertTo-SecureString -String $smtpPassword -AsPlainText -Force
	$smtpCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $smtpUsername, $smtpPWord
}

# Squawk and die if we don't have SQLite
if (!(Test-Path -Path $sqLitePath)) {
	Write-Host -ForegroundColor "Red" -Object "FATAL: Could not find '$sqLitePath'"
	exit
}

# Squawk and die if we don't have a db file
if (!(Test-Path -Path $appStoreDbPath)) {
	Write-Host -ForegroundColor "Red" -Object "FATAL: Could not find '$appStoreDbPath'"
	exit
}

# Squawk and die if we don't have PDQ Deploy executable
if (!(Test-Path -Path $pdqDeployExePath)) {
	Write-Host -ForegroundColor "Red" -Object "FATAL: Could not find '$pdqDeployExePath'"
	exit
}

#
# END INITIALIZE
###################################################################################
###################################################################################
# FUNCTIONS
#

function checkForDbLock () {
	
	$ErrorActionPreference = "SilentlyContinue"
	
	$objFile = New-Object -TypeName System.IO.FileInfo -ArgumentList $appStoreDbPath
	
	[System.IO.FileStream] $fs = $objFile.OpenWrite()
	
	if (!$?) {
		
		Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Database is locked"
		sleep 1
		checkForDbLock
		
	} else {
		
		$fs.Dispose()
		
	}
	
	$ErrorActionPreference = "Continue"
	
}

function myLog ([string] $logMessage) {
	
	[string] $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

	# Available colors are:
	# Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow
	# Gray, DarkGray, Blue, Green, Cyan, Red, Magenta, Yellow, White
	
	# Color Schema:
	# Red = Severe alert (something failed)
	# Yellow = Important alert (changing db settings, etc)
	# White = Important info (value changes, etc)
	# Gray = System info (system activity, etc)
	# Green = Success info (lookup success, etc)
	
	if ($logToConsole -eq $True) {
		Write-Host -Object "$timeStamp - $logMessage"
	}
	
	$logMessage = $logMessage -replace '"','""'
	$logMessage = $logMessage -replace "'","''"

	checkForDbLock
	
	& $sqLitePath $appStoreDbPath "INSERT INTO Log VALUES (Null, '$($Env:ComputerName)', '$timeStamp', '$logMessage')"

}

function sendDeploymentEmailNotification ([string] $userName, [string] $smtpBody) {
	
	$smtpRecipient = $userName + '@domain.com'
	$smtpSubject = "PDQ AppStore deployment notification"

	if ($smtpUseAuth) {
		Send-MailMessage -From $smtpMailFrom -To $smtpRecipient -Subject $smtpSubject -Body $smtpBody -SmtpServer $smtpServer -Credential $smtpCredential -UseSsl
	} else {
		Send-MailMessage -From $smtpMailFrom -To $smtpRecipient -Subject $smtpSubject -Body $smtpBody -SmtpServer $smtpServer
	}
	
}

function sendAdminEmailNotification ([string] $userName, [string] $emailMessage) {

	$smtpRecipient = $userName + '@domain.com'
	$smtpSubject = "Imported PDQ packages into AppStore"
	$smtpBody = $emailMessage

	if ($smtpUseAuth) {
		Send-MailMessage -From $smtpMailFrom -To $smtpRecipient -Subject $smtpSubject -Body $smtpBody -SmtpServer $smtpServer -Credential $smtpCredential -UseSsl
	} else {
		Send-MailMessage -From $smtpMailFrom -To $smtpRecipient -Subject $smtpSubject -Body $smtpBody -SmtpServer $smtpServer
	}
	
}

function deleteDeploymentRow ([int] $queueID) {
	
	checkForDbLock
	
	$rawSql = & $sqLitePath $appStoreDbPath "SELECT PackageName,ComputerName FROM Deployments INNER JOIN Packages ON Packages.PackageID = Deployments.PackageID WHERE QueueID = $queueID"
	
	($packageName, $computerName) = $rawSql.Split('|')
	
	& $sqLitePath $appStoreDbPath "DELETE FROM Deployments WHERE QueueID = $queueID"
	
	#myLog "Server: Deleted deployment for '$PackageName' for computer '$computerName' from database"
}

function populatePackages ([string] $userName) {
	
	[array] $rawPdqPackages = & $pdqDeployExePath GetPackageNames
	
	[int] $importCount = 0
	[int] $duplicateCount = 0
	[int] $invalidCount = 0
	
	:LOOP foreach ($pdqPackageName in $rawPdqPackages) {
		
		[int] $categoryID = 0
		[string] $categoryName = 'Null'
		[int] $packageID = 0
		[string] $packageName = 'Null'
		
		if ($pdqPackageName -match '^\[AppStore\] \((\w+)\) (.*)') {
			
			$categoryName = $Matches[1]
			$packageName = $Matches[2]
			
		} elseif ($pdqPackageName -match '^\[AppStore\] (.*)') {
			
			$packageName = $Matches[1]
			
		} else {
			
			$invalidCount++
			Continue LOOP
			
		}
		
		$categoryName = $categoryName -replace "'","''"
		$categoryName = $categoryName -replace '"','""'
		
		checkForDbLock
		
		if ($categoryName -ne 'Null') {
			
			$categoryID = & $sqLitePath $appStoreDbPath "SELECT CategoryID FROM Categories WHERE CategoryName = '$categoryName'"
			
			if (($categoryID -eq 0) -and ($categoryName.ToLower() -ne 'all')) {
				
				# Insert the category because it doesn't already exist
				myLog "Server: Creating new category '$categoryName'"
				& $sqLitePath $appStoreDbPath "INSERT INTO Categories VALUES (Null, '$categoryName')"
				$categoryID = & $sqLitePath $appStoreDbPath "SELECT CategoryID FROM Categories WHERE CategoryName = '$categoryName'"
				
			} elseif ($categoryName.ToLower() -eq 'all') {
				
				# 'All' category name is reserved and should not be assigned to any packages
				$categoryName = 'Null'
				
			}
			
		}
		
		$packageName = $packageName -replace "'","''"
		$packageName = $packageName -replace '"','""'
		
		checkForDbLock
		
		if ($packageName -ne 'Null') {
			
			$packageID = & $sqLitePath $appStoreDbPath "SELECT PackageID FROM Packages WHERE PackageName = '$packageName'"
			
			if (($packageID -eq 0) -and ($packageName.ToLower() -ne 'all')) {
				
				# Insert the package because it doesn't already exist
				myLog "Server: Inserting package '$packageName' into database"
				& $sqLitePath $appStoreDbPath "INSERT INTO Packages VALUES (Null, '$packageName', '$pdqPackageName', $categoryID, 0, Null, Null)"
				$importCount++
		
			} elseif ($packageName.ToLower() -eq 'all') {
				
				#myLog "Server: Package name 'all' is not allowed; Not importing!"
				$invalidCount++
				
			} else {
				
				#myLog "Server: Not inserting package '$packageName' because it already exists in database"
				$duplicateCount++
				
			}
		
		}
		
	}
	
	if ($sendEmailNotifications -eq $True) {
		sendAdminEmailNotification $userName "Finished importing package list from PDQ.`r`n`r`nImported: $importCount`r`nDuplicates (not imported): $duplicateCount`r`nInvalid (not imported): $invalidCount"
	}
	
}

function deployPackage ([int] $PackageID, [string] $ComputerName, [string] $UserName, [int] $IsScheduled, [string] $ScheduledTime) {
	
	if ($IsScheduled -eq 1) {

		# Get date object
		$sqlDate = Get-Date -Date $ScheduledTime
		
		if (($IsScheduled -eq 1) -and ((Get-Date).Subtract($sqlDate).TotalMilliseconds -le 0)) {
			
			# It's not yet time to deploy
			return(1)
		}
		
	} else {
		
		$ScheduledTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
		
	}
	
	checkForDbLock
	
	$rawSql = & $sqLitePath $appStoreDbPath "SELECT PackageName,PdqPackageName FROM Packages WHERE PackageID = $PackageID"
	
	($packageName, $pdqPackageName) = $rawSql.Split('|')
	
	if ((Test-Connection -ComputerName $ComputerName -Protocol WSMan -BufferSize 1430 -Count 1 -ErrorAction SilentlyContinue).StatusCode -ne 0) { 
		sendAdminEmailNotification $UserName "Package '$packageName' was scheduled for deployment to '$ComputerName' at '$ScheduledTime'.`r`nHowever, the destination computer was not up at the time of deployment.`r`n'$PackageName' was NOT deployed."
		return(0)
	}
	
	if ($pdqPackageName) {
		
		[string] $pdqDeployOutput = & $pdqDeployExePath Deploy -Package "$pdqPackageName" -Targets "$ComputerName"
		
		if (!$pdqDeployOutput) {
			myLog "Server: There was a problem with deploying '$pdqPackageName' to '$ComputerName'"
			return(0)	
		} elseif ($pdqDeployOutput -match 'Package not found') {
			myLog "Server: Package '$pdqPackageName' not found"
			return(0)
		}
		
		if ($pdqDeployOutput -match '^.*ID\s*:\s(\d+)\s?Package\s?:\s(.*)\sTargets') {
			
			$pdqDeploymentID = $Matches[1]
			$pdqFullPackageName = $Matches[2]
			
			myLog "Server: Deploying '$pdqFullPackageName' with PDQ Deploy ID of $pdqDeploymentID to '$ComputerName'"
			
			if (($IsScheduled -eq 1) -and ($sendEmailNotifications -eq $True)) {
				$SmtpBody = "Package '$packageName' which was scheduled for '$ScheduledTime' by '$UserName' has begun deploying to '$ComputerName'.`r`n`r`nIf you have any questions contact IT."
				sendDeploymentEmailNotification $UserName $SmtpBody
			}
			
		} else {
			
			myLog "Server: Unrecognized PDQ Deploy output '$pdqDeployOutput'"
			
		}
		
	}	
	
	return(0)
	
}

function flushDnsCache () {
	
	$output = Start-Process -FilePath 'C:\Windows\System32\ipconfig.exe' -ArgumentList '/flushdns' -PassThru -WindowStyle Hidden
	
}

#
# END FUNCTIONS
###################################################################################
###################################################################################
# MAIN

Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Starting up"

while (1) {
	
	checkForDbLock
	
	flushDnsCache
	
	[array] $rawSqlDeployments = & $sqLitePath $appStoreDbPath "SELECT * FROM Deployments ORDER BY QueueID ASC"
	
	foreach ($rawSqlRow in $rawSqlDeployments) {
		
		($QueueID, $ComputerName, $UserName, $PackageID, $PdqDeploymentID, $IsScheduled, $ScheduledTime) = $rawSqlRow.Split('|')
		
		# Populate database with packages
		if (($PackageID -eq 0) -or ($QueueID -eq 0)) {
						
			populatePackages $UserName
			deleteDeploymentRow $QueueID
			
		# Deploy package to user computer
		} elseif ($PackageID -gt 0) {
			
			$status = deployPackage $PackageID $ComputerName $UserName $IsScheduled $ScheduledTime
			
			if ($status -eq 0) {
				deleteDeploymentRow $QueueID
			}
			
		}
			
	}
	
	sleep 10
}

#
# END MAIN
###################################################################################
