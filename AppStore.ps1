###################################################################################
#
# PDQ AppStore script
#
###################################################################################
# CHANGE LOG
###################################################################################
# 
#
###################################################################################
###################################################################################
# BEGIN VARIABLES
#

# Network SQLite executable location
$sqLiteNetPath = '\\server\share\sqlite3.exe'

# Local SQLite executable location
$sqLiteLocalPath = 'C:\Program Files\SQLite\sqlite3.exe'

# AppStore Network DB Path
$appStoreNetDbPath = '\\server\share\database\as.db'

# Local temp folder
$tempPath = "C:\temp\as"

# AD group members with admin controls
$adminControlsGroup = "acc_IT"

# Log dump folder
$logDumpPath = "C:\temp\logs"

#
# END VARIABLES
###################################################################################

###################################################################################
# BEGIN INITIALIZE
#

# AppStore Local DB Path
$appStoreLocalDbPath = $tempPath + '\as.db'

# Set the preferred SQLite executable location
$sqLitePath = $sqLiteLocalPath

# Set the script name
$scriptName = $MyInvocation.MyCommand.Name

# Set the full script path
$scriptPath = $MyInvocation.MyCommand.Path

# Get AD group memberships
$groupMemberships = ([adsisearcher]"samaccountname=$($Env:USERNAME)").Findone().Properties.memberof -replace '^CN=([^,]+).+$','$1'

# Icon file path
$icoFile = $scriptPath + ".ico"

# Configuration file path
$cfgFile = $scriptPath + ".cfg"


# Squawk and die if we don't have SQLite
if (Test-Path -Path $sqLiteLocalPath) {
	$sqLitePath = $sqLiteLocalPath
} elseif (Test-Path -Path $sqLiteNetPath) {
	$sqLitePath = $sqLiteNetPath
} else {
	Write-Host -ForegroundColor "Red" -Object "FATAL: This script requires sqlite3.exe"
	exit
}

# Squawk and die if we don't have a db file
if (!(Test-Path -Path $appStoreNetDbPath)) {
	Write-Host -ForegroundColor"Red" -Object "FATAL: AppStore database file not found"
	exit
}

# Create the temp folder path or delete all temp files
if (!(Test-Path -Path $tempPath)) {
	New-Item -Type Directory -Path $tempPath
} else {
	Remove-Item -Path "$tempPath\*.astmp" -Force
}

# Copy the icon and config files to the temp folder
if (Test-Path -Path $icoFile) {
	Copy-Item -Path $icoFile -Destination $tempPath | Out-Null
	$icoFile = $tempPath + "\" + $scriptName + ".ico"
}

$localCfgFile = $tempPath + "\" + $scriptName + '.cfg'

if ((Test-Path -Path $cfgFile) -and (!(Test-Path -Path $localCfgFile))) {
	Copy-Item -Path $cfgFile -Destination $tempPath | Out-Null
}

if (Test-Path -Path $localCfgFile) {
	$cfgFile = $localCfgFile
}

# Set up Global configuration hash variable with defaults
$Cfg = @{ 'EnableConsoleLogging' = "Yes";
		  'AlwaysOnTop' = "No" }

#
# END INITIALIZE
###################################################################################

###################################################################################
# BEGIN FUNCTIONS
#

function checkForDbLock () {
	
	$ErrorActionPreference = "SilentlyContinue"
	
	$objFile = New-Object -TypeName System.IO.FileInfo -ArgumentList $appStoreNetDbPath
	
	[System.IO.FileStream] $fs = $objFile.OpenWrite()
	
	if (!$?) {
		
		sleep 1
		checkForDbLock
		
	} else {
		
		$fs.Dispose()
		
	}
	
	$ErrorActionPreference = "Continue"
	
}

function myLog ([string] $logMessage) {
	
	[string] $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	
	if ($Cfg.Get_Item("EnableConsoleLogging") -match "^y") {
	
		# Available colors are:
		# Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow
		# Gray, DarkGray, Blue, Green, Cyan, Red, Magenta, Yellow, White
		
		# Color Schema:
		# Red = Severe alert (something failed)
		# Yellow = Important alert (changing db settings, etc)
		# White = Important info (value changes, etc)
		# Gray = System info (system activity, etc)
		# Green = Success info (lookup success, etc)
		
		Write-Host "$TimeStamp - $logMessage"
		
	}
	
	$logMessage = $logMessage -replace '"','""'
	$logMessage = $logMessage -replace "'","''"
	
	checkForDbLock
	
	& $sqLitePath $appStoreNetDbPath "INSERT INTO Log VALUES (Null, '$($Env:ComputerName)', '$TimeStamp', '$logMessage')"

}

function copyDbToLocal () {
	
	checkForDbLock
	
	Copy-Item -Path "$($appStoreNetDbPath)" -Destination "$($appStoreLocalDbPath)" -Force | Out-Null
	
}

function readConfig () {
	
	if (Test-Path -Path $cfgFile) {
	
		$varTypes = @{ "EnableConsoleLogging" = "bool";
					   "AlwaysOnTop" = "bool" }
						 
		Get-Content $cfgFile | Foreach-Object {

			if ($_ -match "^(\w+)\s*\=\s*(.*)") {
				($Param, $Value) = ($Matches[1], $Matches[2])
				# remove leading space from the value
				$Value = $Value -replace "^\s*", ""
				# environment variable replacement logic for string parameter types
				if ($varTypes.Get_Item($Param) -eq "bool") {
					if ($Value -imatch "^([yt0]|on)") {
						$Value = "Yes"
					} elseif ($Value -imatch "^([nf1]|off)") {
						$Value = "No"
					} else {
						Write-Host -ForegroundColor "Yellow" -Object "Invalid boolean value for $Param; Using default value '$($Cfg.Get_Item($Param))'"
					}
				} else {
					Write-Host -ForegroundColor "Yellow" -Object "Warning: Invalid configuration parameter '$Param'"
				}
				# Set param/value in the global configuration hash
				$Cfg.Set_Item($Param, $Value)
				# remove the param from the type hash
				# we will use this to see if any parameters are undefined
				$varTypes.Remove($Param)
			}

		}
		
		# if any elements remain in the type hash then they are undefined
		if ($varTypes.Count -gt 0) {
			# print warnings for all undefined variables
			$varTypes.Keys | Foreach-Object {
				Write-Host -ForegroundColor "Yellow" -Object "Warning: $_ is undefined; Check '$cfgFile'; Using default value of '$($Cfg.Get_Item($_))'"
			}
		}
	} else {
		Write-Host -ForegroundColor "Yellow" -Object "Warning: Could not open config file '$cfgFile'; Using default settings"
	}
	
}

function okBox ([string] $msgText) {
	
	[void][system.windows.forms.messagebox]::Show($msgText)
	
}

function balloonNotification ([string] $msgText) {
	Add-Type -AssemblyName System.Windows.Forms

	If (-NOT $global:balloon) {
		$global:balloon = New-Object System.Windows.Forms.NotifyIcon

		#Mouse double click on icon to dispose
		[void](Register-ObjectEvent -InputObject $balloon -EventName MouseDoubleClick -SourceIdentifier IconClicked -Action {
			#Perform cleanup actions on balloon tip
			Write-Verbose 'Disposing of balloon'
			$global:balloon.dispose()
			Unregister-Event -SourceIdentifier IconClicked
			Remove-Job -Name IconClicked
			Remove-Variable -Name balloon -Scope Global
		})
	}

	#Need an icon for the tray
	$path = Get-Process -id $pid | Select-Object -ExpandProperty Path

	#Extract the icon from the file
	$balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($icoFile)

	#Can only use certain TipIcons: [System.Windows.Forms.ToolTipIcon] | Get-Member -Static -Type Property
	$balloon.BalloonTipIcon  = [System.Windows.Forms.ToolTipIcon]"Info"
	$balloon.BalloonTipTitle = "AppStore"
	$balloon.BalloonTipText  = $msgText
	$balloon.Visible = $true

	#Display the tip and specify in milliseconds on how long balloon will stay visible
	$balloon.ShowBalloonTip(1000)
	
}

function getCategories () {
	
	[array] $rawSql = & $sqLitePath $appStoreLocalDbPath "SELECT CategoryName FROM Categories WHERE CategoryID > 0 AND NOT CategoryName = 'IT' ORDER BY CategoryName ASC"
	
	if ($isAdmin -eq $True) {
		$rawSql += 'IT'
	}
	
	return($rawSql)
	
}

function getPackages ([string] $categoryName) {
	
	# Get a list of enabled packages for the selected category
	[array] $packages = & $sqLitePath $appStoreLocalDbPath "SELECT Packages.PackageName FROM Packages INNER JOIN Categories ON Packages.CategoryID = Categories.CategoryID WHERE Packages.PackageEnabled = 1 AND Categories.CategoryName = '$categoryName' ORDER BY PackageName ASC"
	
	if (!$packages) {
		return
	}

	[void] $mainFormPackageSelection.Items.Clear()
	
	$packages | Foreach-Object {
		[void] $mainFormPackageSelection.Items.Add($_)
	}
	
}

function importPdqPackages () {
	
	checkForDbLock
	
	& $sqLitePath $appStoreNetDbPath "INSERT INTO Deployments VALUES (null,'$($Env:ComputerName)','$($Env:UserName)',0,null,0,null)"
	
	myLog "Import of all PDQ packages has been queued"
	okBox "Import Queued!"
	
}

function changeCategory ([int] $packageID, [string] $packageName, [string] $categoryName) {
	
	$rawSql = & $sqLitePath $appStoreLocalDbPath "SELECT Categories.CategoryID, Packages.PdqPackageName FROM Categories, Packages WHERE Categories.CategoryName = '$categoryName' AND Packages.PackageName = '$packageName'"
	
	([int] $categoryID, [string] $pdqPackageName) = $rawSql.Split('|')
	
	if ($categoryID -gt 0) {
		
		checkForDbLock
		
		& $sqLitePath $appStoreNetDbPath "UPDATE Packages SET CategoryID = $categoryID WHERE PackageID = $packageID"
	
		myLog "Changed package '$packageName' to category '$categoryName'"
		
	} else {
		
		myLog "Could not get an ID for category '$categoryName'; Not updating category for '$packageName'!"
		
	}
	
	if ($pdqPackageName -match "^\[AppStore\] \(.*?\) (.*)$") {
		
		$newPdqPackageName = "[AppStore] ($categoryName) $($Matches[1])"
		
		checkForDbLock
		
		& $sqLitePath $appStoreNetDbPath "UPDATE Packages SET PdqPackageName = '$newPdqPackageName' WHERE PackageID = $packageID"
	
		myLog "Changed PDQ Name for package '$packageName' to '$newPdqPackageName'"
		
	}
	
	copyDbToLocal

}

function togglePackageEnabled ([int] $packageID, [System.Boolean] $isEnabled) {
	
	$packageName = & $sqLitePath $appStoreLocalDbPath "SELECT PackageName FROM Packages WHERE PackageID = $packageID" 
	
	$packageState = "Disabled"
	$dbState = 0
	
	if ($isEnabled) {
		$packageState = "Enabled"
		$dbState = 1
	}
	
	checkForDbLock
	
	& $sqLitePath $appStoreNetDbPath "UPDATE Packages SET PackageEnabled = $dbState WHERE PackageID = $packageID"
	
	myLog "Changed package '$packageName' to '$packageState'"
	
	copyDbToLocal
}

function togglePackageInteractiveOnly ([int] $packageID, [System.Boolean] $isInteractiveOnly) {
	
	$packageName = & $sqLitePath $appStoreLocalDbPath "SELECT PackageName FROM Packages WHERE PackageID = $packageID" 
	
	$interactiveOnlyState = "Disabled"
	$dbState = 0
	
	if ($isInteractiveOnly) {
		$interactiveOnlyState = "Enabled"
		$dbState = 1
	}
	
	& $sqLitePath $appStoreNetDbPath "UPDATE Packages SET InteractiveOnly = $dbState WHERE PackageID = $packageID"
	
	myLog "Changed package '$packageName' Intactive Only flag to '$interactiveOnlyState'"
	
	copyDbToLocal
}

function deletePackage ([System.Windows.Forms.Button] $deletePackageButton) {
	
	[int] $packageID = $deletePackageButton.Tag
	
	[string] $rawSql = $packageName = & $sqLitePath $appStoreLocalDbPath "SELECT PackageName, PackageEnabled from Packages WHERE PackageID = $packageID"
	
	($packageName, $packageEnabled) = $rawSql.Split('|')
	
	if (($packageID -gt 0) -and ($packageEnabled -lt 1)) {
		
		[int] $inUse = & $sqLitePath $appStoreLocalDbPath "SELECT QueueID FROM Deployments WHERE PackageID = $packageID LIMIT 1"
		
		if ($inUse -gt 0) {
			
			okBox "Package is currently pending deployment and cannot be deleted!"
			return
			
		} else {
			
			
			checkForDbLock

			& $sqLitePath $appStoreNetDbPath "DELETE FROM Packages WHERE PackageID = $packageID"
			
			$deletePackageButton.Dispose()
			
			myLog "'$($Env:UserName)' deleted package '$packageName' from database"
			
		}
		
	} else {
		
		okBox "Package '$packageName' is currently enabled!"
		return
		
	}
	
	copyDbToLocal
}

function editPackages () {
	
	[array] $packagesSqlRaw = & $sqLitePath $appStoreLocalDbPath "SELECT PackageID, PackageName, CategoryID, PackageEnabled, InteractiveOnly FROM Packages WHERE PackageID > 0 ORDER BY PackageName ASC"

	[array] $categoriesSqlRaw = & $sqLitePath $appStoreLocalDbPath "SELECT CategoryID, CategoryName FROM Categories WHERE CategoryID > 0"
	
	$categoryHash = @{}
	
	$categories = @()
	
	foreach ($categoryRow in $categoriesSqlRaw) {
		
		([int] $cid, [string] $cname) = $categoryRow.Split('|')
		
		$categoryHash.Set_Item($cid, $cname)
		
		$categories += $cname
		
	}
	
	$packageEnabledCheckBox = @{}
	$packageInteractiveOnlyCheckBox = @{}
	$packageLabel = @{}
	$categoryDropDown = @{}
	$packageDeleteButton = @{}
	
	$rowHeight = 31
	
	$editForm = New-Object Windows.Forms.Form
	$editForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($icoFile)
	$editForm.AutoScroll = $True
	$editForm.AutoSize = $True
	$editForm.Text = "Edit Packages"
	$editForm.AutoSizeMode = "GrowAndShrink"
	$editForm.StartPosition = "CenterScreen"
	$editForm.MaximumSize = New-Object System.Drawing.Size(1200,700)
	$editForm.Opacity = 1
	$editForm.KeyPreview = $False
	# Font styles are: Regular, Bold, Italic, Underline, Strikeout
	$editForm.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
	
	$y = 0
	foreach ($row in $packagesSqlRaw) {
		([int] $PackageID, [string] $PackageName, [int] $CategoryID, [int] $PackageEnabled, [int] $InteractiveOnly) = $row.Split('|')
		
		if ($CategoryID -match '\d') {
			$CategoryName = $categoryHash.$CategoryID
		} else {
			$CategoryName = $Null
		}
		
		# Package enable checkbox
		$packageEnabledCheckBox[$PackageID] = New-Object System.Windows.Forms.Checkbox
		$packageEnabledCheckBox[$PackageID].Text = "Enabled"
		$packageEnabledCheckBox[$PackageID].Name = $PackageID
		if ($PackageEnabled -eq 1) {
			$packageEnabledCheckBox[$PackageID].Checked = $True
		} else {
			$packageEnabledCheckBox[$PackageID].Checked = $False
		}
		$packageEnabledCheckBox[$PackageID].Size = New-Object System.Drawing.Size(85,30)
		$packageEnabledCheckBox[$PackageID].Location = New-Object System.Drawing.Point(5,$y)
		$packageEnabledCheckBox[$PackageID].Add_CheckedChanged({ togglePackageEnabled $($this.Name) $($this.Checked) })
		$editForm.Controls.Add($packageEnabledCheckBox[$PackageID])
		
		# Package interactive only checkbox
		$packageInteractiveOnlyCheckBox[$PackageID] = New-Object System.Windows.Forms.Checkbox
		$packageInteractiveOnlyCheckBox[$PackageID].Text = "Interactive Only"
		$packageInteractiveOnlyCheckBox[$PackageID].TextAlign = 'MiddleCenter'
		$packageInteractiveOnlyCheckBox[$PackageID].Name = $PackageID
		if ($InteractiveOnly -eq 1) {
			$packageInteractiveOnlyCheckBox[$PackageID].Checked = $True
		} else {
			$packageInteractiveOnlyCheckBox[$PackageID].Checked = $False
		}
		$packageInteractiveOnlyCheckBox[$PackageID].Size = New-Object System.Drawing.Size(140,30)
		$packageInteractiveOnlyCheckBox[$PackageID].Location = New-Object System.Drawing.Point(($($packageEnabledCheckBox[$PackageID]).Right + 5),$y)
		$packageInteractiveOnlyCheckBox[$PackageID].Add_CheckedChanged({ togglePackageInteractiveOnly $($this.Name) $($this.Checked) })
		$editForm.Controls.Add($packageInteractiveOnlyCheckBox[$PackageID])
		
		# Drop down for the category
		$categoryDropDown[$PackageID] = New-Object System.Windows.Forms.Combobox
		$categoryDropDown[$PackageID].Name = $PackageID
		$categoryDropDown[$PackageID].Text = $CategoryName
		$categoryDropDown[$PackageID].Tag = $PackageName
		$categories | Foreach-Object {
			$categoryDropDown[$PackageID].Items.Add($_)
		}
		$categoryDropDown[$PackageID].Size = New-Object System.Drawing.Size(100,30)
		$categoryDropDown[$PackageID].Location = New-Object System.Drawing.Point(($($packageInteractiveOnlyCheckBox[$PackageID]).Right + 5),($y + 2))
		$categoryDropDown[$PackageID].Add_SelectedValueChanged({ changeCategory $($this.Name) $($this.Tag) $($this.Text) })
		$editForm.Controls.Add($categoryDropDown[$PackageID])
		
		# Package delete button
		$packageDeleteButton[$PackageID] = New-Object System.Windows.Forms.Button
		$packageDeleteButton[$PackageID].Tag = $PackageID
		$packageDeleteButton[$PackageID].Text = "Delete"
		$packageDeleteButton[$PackageID].TextAlign = 'MiddleCenter'
		$packageDeleteButton[$PackageID].Size = New-Object System.Drawing.Size(90,30)
		$packageDeleteButton[$PackageID].Location = New-Object System.Drawing.Size(($($categoryDropDown[$PackageID]).Right + 5),$y)
		$packageDeleteButton[$PackageID].Add_Click({ deletePackage $this }) 
		$editForm.Controls.Add($packageDeleteButton[$PackageID])
		
		# Package label
		$packageLabel[$PackageID] = New-Object System.Windows.Forms.Label
		$packageLabel[$PackageID].Text = $PackageName
		$packageLabel[$PackageID].TextAlign = 'MiddleLeft'
		$packageLabel[$PackageID].Size = New-Object System.Drawing.Size(600,30)
		$packageLabel[$PackageID].Location = New-Object System.Drawing.Point(($($packageDeleteButton[$PackageID]).Right + 5),$y)
		$editForm.Controls.Add($packageLabel[$PackageID])	
		
		$y += $rowHeight
		
	}
	
	if ($Cfg.Get_Item("AlwaysOnTop") -imatch "^y") {
		$editForm.Topmost = $True
	} else {
		$editForm.Topmost = $False
	}
	
	[void]$editForm.ShowDialog()
	
}

function deleteDeployment ([System.Windows.Forms.Button] $PackageDeleteButton) {
	
	[int] $QueueID = $PackageDeleteButton.Tag
	
	$rawSql = & $sqLitePath $appStoreLocalDbPath "SELECT PackageName, ScheduledTime FROM Deployments INNER JOIN Packages ON Deployments.PackageID = Packages.PackageID WHERE Deployments.QueueID = $QueueID"
	
	($PackageName, $ScheduledTime) = $rawSql.Split('|')
	
	checkForDbLock
	
	& $sqLitePath $appStoreNetDbPath "DELETE FROM Deployments WHERE QueueID = $QueueID"
	
	okBox "Deployment for '$PackageName' at '$ScheduledTime' has been deleted!"
	
	$PackageDeleteButton.Enabled = $False
	
	copyDbToLocal
	
}

function viewDeployments {
	
	$rawSql = $Null
	if ($isAdmin -eq $True) {
		$rawSql = & $sqLitePath $appStoreLocalDbPath "SELECT QueueID, ComputerName, UserName, PackageName, ScheduledTime FROM Deployments INNER JOIN Packages ON Deployments.PackageID = Packages.PackageID"
	} else {
		$rawSql = & $sqLitePath $appStoreLocalDbPath "SELECT QueueID, ComputerName, UserName, PackageName, ScheduledTime FROM Deployments INNER JOIN Packages ON Deployments.PackageID = Packages.PackageID WHERE Deployments.ComputerName = '$($Env:ComputerName)'"
	}
	
	if (!$rawSql) {
		if ($isAdmin -eq $True) {
			okBox "There are no pending deployments in the install queue."
		} else {
			okBox "There are no pending deployments scheduled for this computer.`n`nIf you just deployed software, the installation is already underway.`n`nIf you deployed software and never received a pop-up message, then the deployment likely failed; Either try deploying again or contact IT."
		}
		return
	}
	
	$packageLabel = @{}
	$packageDeleteButton = @{}
	
	$rowHeight = 31
	
	$viewDeploymentsForm = New-Object Windows.Forms.Form
	$viewDeploymentsForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($icoFile)
	$viewDeploymentsForm.AutoScroll = $True
	$viewDeploymentsForm.AutoSize = $True
	$viewDeploymentsForm.Text = "Pending Deployments"
	$viewDeploymentsForm.AutoSizeMode = "GrowAndShrink"
	$viewDeploymentsForm.StartPosition = "CenterScreen"
	$viewDeploymentsForm.MaximumSize = New-Object System.Drawing.Size(1500,700)
	$viewDeploymentsForm.Opacity = 1
	$viewDeploymentsForm.KeyPreview = $False
	# Font styles are: Regular, Bold, Italic, Underline, Strikeout
	$viewDeploymentsForm.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
	
	$y = 0
	foreach ($row in $rawSql) {
		([int] $QueueID, [string] $ComputerName, [string] $UserName, [string] $PackageName, [string] $ScheduledTime) = $row.Split('|')
		
		if ($ScheduledTime -notmatch "^\d") {
			$ScheduledTime = Get-Date -Format "MM/dd/yyyy HH:mm"
		}
		
		# Package delete button
		$packageDeleteButton[$QueueID] = New-Object System.Windows.Forms.Button
		$packageDeleteButton[$QueueID].Tag = $QueueID
		$packageDeleteButton[$QueueID].Text = "Delete Deployment"
		$packageDeleteButton[$QueueID].TextAlign = 'MiddleCenter'
		$packageDeleteButton[$QueueID].Size = New-Object System.Drawing.Size(200,30)
		$packageDeleteButton[$QueueID].Location = New-Object System.Drawing.Size(5,$y)
		$packageDeleteButton[$QueueID].Add_Click({ deleteDeployment $this }) 
		$viewDeploymentsForm.Controls.Add($packageDeleteButton[$QueueID])
		
		# Package label
		$packageLabel[$QueueID] = New-Object System.Windows.Forms.Label
		$packageLabel[$QueueID].Text = "'$PackageName' was queued by '$UserName' for deployment to '$ComputerName' at '$ScheduledTime'"
		$packageLabel[$QueueID].TextAlign = 'MiddleLeft'
		$packageLabel[$QueueID].Size = New-Object System.Drawing.Size(900,30)
		$packageLabel[$QueueID].Location = New-Object System.Drawing.Point(($($packageDeleteButton[$QueueID]).Right + 5),$y)
		$viewDeploymentsForm.Controls.Add($packageLabel[$QueueID])	
		
		$y += $rowHeight
		
	}
	
	if ($Cfg.Get_Item("AlwaysOnTop") -imatch "^y") {
		$viewDeploymentsForm.Topmost = $True
	} else {
		$viewDeploymentsForm.Topmost = $False
	}
	
	[void]$viewDeploymentsForm.ShowDialog()
	
}

function viewDeploymentHistory ([string] $PackageName) {

	if (!$PackageName) {
		okBox "Nothing selected!"
		return
	}
	
	$rawSql = $Null
	if ($isAdmin -eq $True) {
		$rawSql = & $sqLitePath $appStoreLocalDbPath "SELECT LogID, TimeStamp, Message FROM Log WHERE Message LIKE '% $PackageName'' with PDQ Deploy ID of %'"
	} else {
		$rawSql = & $sqLitePath $appStoreLocalDbPath "SELECT LogID, TimeStamp, Message FROM Log WHERE Message LIKE '% $PackageName'' with PDQ Deploy ID of % to ''$($Env:ComputerName)'''"
	}
	
	if (!$rawSql) {
		if ($isAdmin -eq $True) {
			okBox "'$PackageName' has never been deployed from AppStore to any computer."
		} else {
			okBox "'$PackageName' has never been deployed from AppStore to this computer."
		}
		return
	}
	
	$timeStampLabel = @{}
	$messageLabel = @{}
	
	$rowHeight = 31
	
	$viewDeploymentHistoryForm = New-Object Windows.Forms.Form
	$viewDeploymentHistoryForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($icoFile)
	$viewDeploymentHistoryForm.AutoScroll = $True
	$viewDeploymentHistoryForm.AutoSize = $True
	if ($isAdmin -eq $True) {
		$viewDeploymentHistoryForm.Text = "Deployment History for '$PackageName'"
	} else {
		$viewDeploymentHistoryForm.Text = "Deployment History for '$PackageName' on $($Env:ComputerName)"
	}
	$viewDeploymentHistoryForm.AutoSizeMode = "GrowAndShrink"
	$viewDeploymentHistoryForm.StartPosition = "CenterScreen"
	$viewDeploymentHistoryForm.MaximumSize = New-Object System.Drawing.Size(1500,700)
	$viewDeploymentHistoryForm.Opacity = 1
	$viewDeploymentHistoryForm.KeyPreview = $False
	# Font styles are: Regular, Bold, Italic, Underline, Strikeout
	$viewDeploymentHistoryForm.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
	
	$y = 0
	foreach ($row in $rawSql) {
		([int] $LogID, [string] $TimeStamp, [string] $LogMessage) = $row.Split('|')
		
		$MessageLabelText = $LogMessage
		if ($LogMessage -match "Deploy ID of (\d+) to '(\S+)'") {
			$MessageLabelText = "'$PackageName' was deployed to '$($Matches[2])' with ID '$($Matches[1])'"
		}
		
		# Package delete button
		$timeStampLabel[$LogID] = New-Object System.Windows.Forms.Label
		$timeStampLabel[$LogID].Text = $TimeStamp
		$timeStampLabel[$LogID].TextAlign = 'MiddleCenter'
		$timeStampLabel[$LogID].Size = New-Object System.Drawing.Size(200,30)
		$timeStampLabel[$LogID].Location = New-Object System.Drawing.Size(5,$y)
		$viewDeploymentHistoryForm.Controls.Add($timeStampLabel[$LogID])
		
		# Package label
		$messageLabel[$LogID] = New-Object System.Windows.Forms.Label
		$messageLabel[$LogID].Text = $MessageLabelText
		$messageLabel[$LogID].TextAlign = 'MiddleLeft'
		$messageLabel[$LogID].Size = New-Object System.Drawing.Size(800,30)
		$messageLabel[$LogID].Location = New-Object System.Drawing.Point(($($timeStampLabel[$LogID]).Right + 5),$y)
		$viewDeploymentHistoryForm.Controls.Add($messageLabel[$LogID])	
		
		$y += $rowHeight
		
	}
	
	if ($Cfg.Get_Item("AlwaysOnTop") -imatch "^y") {
		$viewDeploymentHistoryForm.Topmost = $True
	} else {
		$viewDeploymentHistoryForm.Topmost = $False
	}
	
	[void]$viewDeploymentHistoryForm.ShowDialog()
}

function deleteDeployments ([string] $deleteScope) {
	
	checkForDbLock
	
	if (($deleteScope -eq "All") -and ($isAdmin -eq $True)) {
		
		& $sqLitePath $appStoreNetDbPath "DELETE FROM Deployments"
	
		myLog "Deleted all pending deployments from database"
		okBox "All pending deployments have been deleted!"
		
	} elseif ($deleteScope -eq "Computer") {
		
		& $sqLitePath $appStoreNetDbPath "DELETE FROM Deployments WHERE ComputerName = '$($Env:ComputerName)'"
	
		myLog "Deleted deployments for '$($Env:ComputerName)'"
		okBox "Pending deployments for '$($Env:ComputerName)' have been deleted!"

	}
	
	copyDbToLocal
}

function createCategory ([System.Windows.Forms.TextBox] $categoryTextBox) {
	
	$categoryName = $categoryTextBox.Text
	
	if (!$categoryName) {
		okBox "No category given!"
		return
	}
	
	$categoryName = $categoryName -replace "'","''"
	$categoryName = $categoryName -replace '"','""'
	
	checkForDbLock
	
	& $sqLitePath $appStoreNetDbPath "INSERT INTO Categories VALUES (null, '$categoryName')"
	
	[array] $categories = getCategories
	
	$mainFormCategoryDropDown.DataSource = $categories
	[void] $mainFormCategoryDropDown.Refresh()
	
	myLog "'$($Env:UserName)' added category '$categoryName' to the database"
	okBox "Added '$categoryName'"	
	
	copyDbToLocal
}

function createCategoryInput () {
	
	$createCategoryForm = New-Object Windows.Forms.Form
	$createCategoryForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($icoFile)
	$createCategoryForm.AutoScroll = $True
	$createCategoryForm.AutoSize = $True
	$createCategoryForm.Text = "Create Category Name"
	$createCategoryForm.AutoSizeMode = "GrowAndShrink"
	$createCategoryForm.StartPosition = "CenterScreen"
	$createCategoryForm.Opacity = 1
	$createCategoryForm.KeyPreview = $False
	# Font styles are: Regular, Bold, Italic, Underline, Strikeout
	$createCategoryForm.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
	
	$label1 = New-Object System.Windows.Forms.Label
	$label1.Text = "Category Name:"
	$label1.Size = New-Object System.Drawing.Size(120,30)
	$label1.location = New-Object System.Drawing.Point(10,10)
	$label1.Font = 'Microsoft Sans Serif,10'
	$createCategoryForm.Controls.Add($label1)
	
	$textbox1 = New-Object System.Windows.Forms.TextBox
	$textbox1.Size = New-Object System.Drawing.Size(230,30)
	$textbox1.Location = New-Object System.Drawing.Point(135,10)
	$textbox1.Font = 'Microsoft Sans Serif,10'
	$createCategoryForm.Controls.Add($textbox1) 
	
	$button1 = New-Object system.Windows.Forms.Button
	$button1.text = "Create"
	$button1.Size = New-Object System.Drawing.Point(90,30)
	$button1.location = New-Object System.Drawing.Point(370,5)
	$button1.Font = 'Microsoft Sans Serif,10'
	$button1.Add_Click({ 
		createCategory $textbox1
		$createCategoryForm.Close()
	})
	$createCategoryForm.Controls.Add($button1)
	
	if ($Cfg.Get_Item("AlwaysOnTop") -imatch "^y") {
		$createCategoryForm.Topmost = $True
	} else {
		$createCategoryForm.Topmost = $False
	}
	
	[void] $createCategoryForm.ShowDialog()
	
}

function deleteCategory ([System.Windows.Forms.Combobox] $categoryDropDown) {
	
	$categoryName = $categoryDropDown.Text
	
	if (!$categoryName) {
		okBox "No category selected!"
		return
	} elseif ($categoryName -eq 'IT') {
		okBox "Cannot delete the 'IT' category!"
		return
	}
	
	[int] $inUse = & $sqLitePath $appStoreLocalDbPath "SELECT PackageID from Packages WHERE CategoryID = (SELECT CategoryID FROM Categories WHERE CategoryName = '$categoryName') LIMIT 1"
	
	if ($inUse -gt 0) {
		okBox "Category '$categoryName' is in use!"
		return
	}
	
	checkForDbLock
	
	& $sqLitePath $appStoreNetDbPath "DELETE FROM Categories WHERE CategoryName = '$categoryName'"
	
	[array] $categories = getCategories
	
	$categoryDropDown.DataSource = $categories
	$categoryDropDown.Refresh()
	
	$mainFormCategoryDropDown.DataSource = $categories
	$mainFormCategoryDropDown.Refresh()
	
	myLog "'$($Env:UserName)' deleted category '$categoryName' from the database"
	okBox "Removed '$categoryName'"
	
	copyDbToLocal
}

function deleteCategoryPicker () {
	
	[array] $categories = getCategories

	$categoryDropDown = @{}
	
	$deleteCategoryPickerForm = New-Object Windows.Forms.Form
	$deleteCategoryPickerForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($icoFile)
	$deleteCategoryPickerForm.AutoScroll = $True
	$deleteCategoryPickerForm.AutoSize = $True
	$deleteCategoryPickerForm.Text = "Delete Category"
	$deleteCategoryPickerForm.AutoSizeMode = "GrowAndShrink"
	$deleteCategoryPickerForm.StartPosition = "CenterScreen"
	$deleteCategoryPickerForm.Opacity = 1
	$deleteCategoryPickerForm.KeyPreview = $False
	# Font styles are: Regular, Bold, Italic, Underline, Strikeout
	$deleteCategoryPickerForm.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)

	# Drop down for the category
	$categoryDropDown = New-Object System.Windows.Forms.Combobox
	$categoryDropDown.Text = ""
	$categoryDropDown.DataSource = $categories
	$categoryDropDown.Refresh()
	$categoryDropDown.Size = New-Object System.Drawing.Size(150,30)
	$categoryDropDown.Location = New-Object System.Drawing.Point(5,5)
	$deleteCategoryPickerForm.Controls.Add($categoryDropDown)
	
	# Submit button
	$submitButton = New-Object System.Windows.Forms.Button
	$submitButton.Text = "Submit"
	$submitButton.Size = New-Object System.Drawing.Size(90,30)
	$submitButton.Location = New-Object System.Drawing.Point(155,5)
	$submitButton.Add_Click({ 
		deleteCategory $categoryDropDown
	})
	$deleteCategoryPickerForm.Controls.Add($submitButton)	
		
	if ($Cfg.Get_Item("AlwaysOnTop") -imatch "^y") {
		$deleteCategoryPickerForm.Topmost = $True
	} else {
		$deleteCategoryPickerForm.Topmost = $False
	}
	
	[void]$deleteCategoryPickerForm.ShowDialog()

}

function dumpAppStoreLog () {
	
	[array] $logMessages = & $sqLitePath $appStoreLocalDbPath "SELECT ComputerName, TimeStamp, Message FROM Log"
	
	if (!(Test-Path -Path $logDumpPath)) {
		New-Item -Path $logDumpPath -ItemType Directory | Out-Null
	}
	
	$filePath = $logDumpPath + "\" + (Get-Date).Ticks + "_AppStore_Log.csv"
	
	"SEP=,`nTime Stamp,Computer Name,Log Message" | Out-File -FilePath $filePath
	
	foreach ($Entry in $logMessages) {
		($ComputerName, $TimeStamp, $Message) = $Entry.Split('|')
		"$TimeStamp,$ComputerName,$Message" | Out-File -FilePath $filePath -Append
	}
	
	myLog "AppStore log dumped by '$($Env:UserName)' on '$($Env:ComputerName)'"
	okBox "Log dumped to '$filePath'"
}

function queueDeployment ([string] $Package) {
	
	if (!$Package) {
		okBox "Nothing selected!"
		return
	}
	
	checkForDbLock
	
	[int] $deploymentID = & $sqLitePath $appStoreLocalDbPath "SELECT QueueID FROM Deployments WHERE PackageID = (SELECT PackageID FROM Packages WHERE PackageName = '$Package') AND ComputerName = '$($Env:ComputerName)'"
	
	[int] $interactiveOnly = 0
	
	if ($deploymentID -gt 0) {
		& $sqLitePath $appStoreNetDbPath "UPDATE Deployments SET IsScheduled = 0, ScheduledTime = Null WHERE QueueID = $deploymentID"
	} else {
		& $sqLitePath $appStoreNetDbPath "INSERT INTO Deployments VALUES (Null, '$($Env:ComputerName)', '$($Env:UserName)', (SELECT PackageID FROM Packages WHERE PackageName = '$Package'), Null, 0, Null)"
		$interactiveOnly = & $sqLitePath $appStoreLocalDbPath "SELECT InteractiveOnly FROM Packages WHERE PackageName = '$Package'"
	}
	
	myLog "Package '$package' is queued for immediate deployment"
	
	if ($interactiveOnly) {
		okBox "Package '$package' will begin installing shortly!`n`nYou may see additional windows and/or pop-up messages as the deployment progresses.`n`nDO NOT REBOOT YOUR COMPUTER UNTIL THE DEPLOYMENT HAS FINISHED!"
	} else {
		okBox "Package '$package' will begin deploying to your computer silently; You will not see any progress bars or other indications of an installation process!`n`nTHIS IS THE LAST MESSAGE YOU WILL SEE UNTIL THE DEPLOYMENT HAS COMPLETED!`n`nDO NOT REBOOT YOUR COMPUTER UNTIL YOU SEE A POP-UP MESSAGE!"
	}
}

function scheduleDeployment ([string] $Package, [string] $HourText, [string] $MinuteText, [bool] $TomorrowChecked ) {
	
	if (($HourText -eq "Hour") -or ($MinuteText -eq "Min")) {
		okBox "Invalid time selected. Try again."
		return
	}
	
	$Time = $HourText + ':' + $MinuteText + ':' + $((Get-Date).Second).ToString().PadLeft(2,'0')
	
	[int] $Hour = $HourText
	[int] $Minute = $MinuteText
	
	$curDate = Get-Date
	
	if (($Hour -le $curDate.Hour) -and ($Minute -le $curDate.Minute) -and ($TomorrowChecked -eq $False)) {
		okBox "!! Selected time is in the past; Try again !!"
		return
	}
	
	$ScheduledTime = Get-Date -Hour $Hour -Minute $Minute
	
	if ($TomorrowChecked -eq $True) {
		$ScheduledTime = (Get-Date -Hour $Hour -Minute $Minute).AddDays(1)
	}
	
	checkForDbLock
	
	# Check for a previous deployment of this package on this computer
	[int] $deploymentID = & $sqLitePath $appStoreLocalDbPath "SELECT QueueID FROM Deployments WHERE PackageID = (SELECT PackageID FROM Packages WHERE PackageName = '$Package') AND ComputerName = '$($Env:ComputerName)'"
	
	if ($deploymentID -gt 0) {
		if ($TomorrowChecked -eq $True) {
			& $sqLitePath $appStoreNetDbPath "UPDATE Deployments SET IsScheduled = 1, ScheduledTime = date('now','localtime','+1 days') || ' ' || '$Time' WHERE QueueID = $deploymentID"
		} else {
			& $sqLitePath $appStoreNetDbPath "UPDATE Deployments SET IsScheduled = 1, ScheduledTime = date('now','localtime') || ' ' || '$Time' WHERE QueueID = $deploymentID"
		}
	} else {
		if ($TomorrowChecked -eq $True) {
			& $sqLitePath $appStoreNetDbPath "INSERT INTO Deployments VALUES (Null, '$($Env:ComputerName)', '$($Env:UserName)', (SELECT PackageID FROM Packages WHERE PackageName = '$Package'), Null, 1, date('now','localtime','+1 days') || ' ' || '$Time')"
		} else {
			& $sqLitePath $appStoreNetDbPath "INSERT INTO Deployments VALUES (Null, '$($Env:ComputerName)', '$($Env:UserName)', (SELECT PackageID FROM Packages WHERE PackageName = '$Package'), Null, 1, date('now','localtime') || ' ' || '$Time')"
		}
	}
	
	myLog "Package '$package' is queued for '$ScheduledTime'"
	okBox "Package '$package' has been queued for '$ScheduledTime'`n`nYou DO NOT need to be logged in at the time of deployment, however, your computer must be online!"
	
}

function schedulePicker ([string] $Package) {
	
	if (!$Package) {
		okBox "Nothing selected!"
		return
	}
	
	[array] $availableHours = '00'
	[array] $availableMinutes = '00'
	
	for ($h=1; $h -le 23; $h++) {
		[string] $tmp = $h
		$availableHours += $tmp.PadLeft(2,"0")	
	}
	
	for ($m=1; $m -le 59; $m++) {
		[string] $tmp = $m
		$availableMinutes += $tmp.PadLeft(2,"0")
	}
	
	$pickerForm = New-Object Windows.Forms.Form
	$pickerForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($icoFile)
	$pickerForm.AutoScroll = $True
	$pickerForm.AutoSize = $True
	$pickerForm.Text = "Scheduler"
	$pickerForm.AutoSizeMode = "GrowAndShrink"
	$pickerForm.StartPosition = "CenterScreen"
	$pickerForm.Opacity = 1
	$pickerForm.KeyPreview = $False
	$pickerFont = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
	# Font styles are: Regular, Bold, Italic, Underline, Strikeout
	$pickerForm.Font = $pickerFont

	# Drop down for the hours
	$hourDropDown = New-Object System.Windows.Forms.Combobox
	$hourDropDown.Text = "Hour"
	$availableHours | Foreach-Object {
		[void] $hourDropDown.Items.Add($_)
	}
	$hourDropDown.Size = New-Object System.Drawing.Size(60,30)
	$hourDropDown.Location = New-Object System.Drawing.Point(5,5)
	$pickerForm.Controls.Add($hourDropDown)
	
	# Drop down for the minutes
	$minuteDropDown = New-Object System.Windows.Forms.Combobox
	$minuteDropDown.Text = "Min"
	$availableMinutes | Foreach-Object {
		[void] $minuteDropDown.Items.Add($_)
	}
	$minuteDropDown.Size = New-Object System.Drawing.Size(60,30)
	$minuteDropDown.Location = New-Object System.Drawing.Point(($hourDropDown.Right + 5),5)
	$pickerForm.Controls.Add($minuteDropDown)
	
	$tomorrowCheckbox = New-Object System.Windows.Forms.Checkbox
	$tomorrowCheckbox.Size = New-Object System.Drawing.Size(13,30)
	$tomorrowCheckbox.Location = New-Object System.Drawing.Point(5,($hourDropDown.Bottom + 5))
	$pickerForm.Controls.Add($tomorrowCheckbox)
	
	$tomorrowLabel = New-Object system.Windows.Forms.Label
	$tomorrowLabel.Text = "Tomorrow"
	$tomorrowLabel.TextAlign = 'MiddleCenter'
	$tomorrowLabel.Size = New-Object System.Drawing.Size(80,30)
	$tomorrowLabel.Location = New-Object System.Drawing.Point(($tomorrowCheckbox.Right + 5),($hourDropDown.Bottom + 5))
	$pickerForm.Controls.Add($tomorrowLabel)
	
	
	$submitSchedule = New-Object system.Windows.Forms.Button
	$submitSchedule.Text = "Submit"
	$submitSchedule.Size = New-Object System.Drawing.Size(70,30)
	$submitSchedule.Location = New-Object System.Drawing.Point(5,($tomorrowCheckbox.Bottom + 10))
	$pickerForm.Controls.Add($submitSchedule)
	
	$submitSchedule.Add_Click({
		$status = scheduleDeployment $Package $hourDropDown.Text $minuteDropDown.Text $tomorrowCheckbox.Checked
		$pickerForm.Close()
	})
	
	if ($Cfg.Get_Item("AlwaysOnTop") -imatch "^y") {
		$pickerForm.Topmost = $True
	} else {
		$pickerForm.Topmost = $False
	}
	
	[void] $pickerForm.ShowDialog()
	
}

function showAdminOptions () {

	$adminForm = New-Object system.Windows.Forms.Form
	$adminForm.ClientSize = '220,270'
	$adminForm.text = "Admin"
	$adminForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($icoFile)
	$adminForm.StartPosition = "CenterScreen"
	
	$button1 = New-Object system.Windows.Forms.Button
	$button1.text = "Import PDQ Packages"
	$button1.Size = New-Object System.Drawing.Size (180,30)
	$button1.location = New-Object System.Drawing.Point(10,20)
	$button1.Font = 'Microsoft Sans Serif,10'
	$button1.Add_Click({ importPdqPackages })

	$button2 = New-Object system.Windows.Forms.Button
	$button2.text = "Create Category"
	$button2.Size = New-Object System.Drawing.Size (180,30)
	$button2.location = New-Object System.Drawing.Point(10,($button1.Bottom + 5))
	$button2.Font = 'Microsoft Sans Serif,10'
	$button2.Add_Click({ createCategoryInput })
	
	$button3 = New-Object system.Windows.Forms.Button
	$button3.text = "Delete Category"
	$button3.Size = New-Object System.Drawing.Size (180,30)
	$button3.location = New-Object System.Drawing.Point(10,($button2.Bottom + 5))
	$button3.Font = 'Microsoft Sans Serif,10'
	$button3.Add_Click({ deleteCategoryPicker })
	
	$button4 = New-Object system.Windows.Forms.Button
	$button4.text = "Edit Packages"
	$button4.Size = New-Object System.Drawing.Size (180,30)
	$button4.location = New-Object System.Drawing.Point(10,($button3.Bottom + 5))
	$button4.Font = 'Microsoft Sans Serif,10'
	$button4.Add_Click({ editPackages })
	
	$button5 = New-Object system.Windows.Forms.Button
	$button5.text = "Delete All Deployments"
	$button5.Size = New-Object System.Drawing.Size (180,30)
	$button5.location = New-Object System.Drawing.Point(10,($button4.Bottom + 5))
	$button5.Font = 'Microsoft Sans Serif,10'
	$button5.Add_Click({ deleteDeployments "All" })

	$button6 = New-Object system.Windows.Forms.Button
	$button6.text = "Dump Log"
	$button6.Size = New-Object System.Drawing.Size (180,30)
	$button6.location = New-Object System.Drawing.Point(10,($button5.Bottom + 5))
	$button6.Font = 'Microsoft Sans Serif,10'
	$button6.Add_Click({ dumpAppStoreLog })

	$groupButtonEvent = New-Object system.Windows.Forms.Groupbox
	$groupButtonEvent.Text = "Admin Controls"
	$groupButtonEvent.Size = New-Object System.Drawing.Size (205,235)
	$groupButtonEvent.location = New-Object System.Drawing.Point(10,5)
	$groupButtonEvent.controls.AddRange(@($button1,$button2,$button3,$button4,$button5,$button6))
	
	$adminForm.controls.Add($groupButtonEvent)
	
	if ($Cfg.Get_Item("AlwaysOnTop") -imatch "^y") {
		$adminForm.Topmost = $True
	} else {
		$adminForm.Topmost = $False
	}
	
	[void]$adminForm.ShowDialog()
	
}

function getPackageDescription ([string] $packageName) {

	# Get the CategoryID for the selected category
	[string] $rawSql = & $sqLitePath $appStoreLocalDbPath "SELECT InteractiveOnly, PackageDescription FROM Packages WHERE PackageName = '$packageName'"

	[int] $interactiveOnly = 0
	[string] $packageDescription = $null
	
	($interactiveOnly, $packageDescription) = $rawSql.Split('|')
	
	$packageDescription = $packageDescription -replace "__CRLF__","`r`n"
	$packageDescription = $packageDescription -replace "__SQ__","'"
	$packageDescription = $packageDescription -replace "__DQ__",'"'
	
	$mainFormPackageDescription.Text = $packageDescription
	
}

function updateDescription ([string] $packageName, [string] $packageDescription) {
	
	if ((!$packageName) -or (!$packageDescription)) {
		okBox "Nothing selected!"
		return
	}
	
	$sanitizedPackageDescription = $packageDescription -replace "'","__SQ__"
	$sanitizedPackageDescription = $sanitizedPackageDescription -replace '"','__DQ__'
	$sanitizedPackageDescription = $sanitizedPackageDescription -replace "`r`n","__CRLF__"
	
	checkForDbLock
	
	$status = & $sqLitePath $appStoreNetDbPath "UPDATE Packages SET PackageDescription = '$sanitizedPackageDescription' WHERE PackageName = '$packageName'"
	
	if ($status) {
		okBox $status
	} else {
		myLog "'$($Env:UserName)' updated description for package '$packageName'"
		okBox "Updated package description!"
	}
	
	copyDbToLocal
}

function checkInteractiveOnly ([string] $packageName) {
	
	[int] $interactiveOnly = 0
	
	$interactiveOnly = & $sqLitePath $appStoreLocalDbPath "SELECT InteractiveOnly FROM Packages WHERE PackageName = '$packageName'"
	
	if ($interactiveOnly -gt 0) {
		$mainFormButton2.Enabled = $False
	}
	
}

function toggleDeploybuttons ([int] $enableDisable, [string] $packageName) {
	
	if (($enableDisable -eq 0) -and ($packageName)) {
		$packageName | Out-File -FilePath "$tempPath\disabledPackages.astmp" -Append
		$mainFormButton1.Enabled = $False
		$mainFormButton2.Enabled = $False
	} elseif (($enableDisable -eq 1) -and ($packageName)) {
		$mainFormButton1.Enabled = $True
		$mainFormButton2.Enabled = $True
		if (Test-Path -Path "$tempPath\disabledPackages.astmp") {
			if ((Get-Content -Path "$tempPath\disabledPackages.astmp").Contains($packageName)) {
				$mainFormButton1.Enabled = $False
				$mainFormButton2.Enabled = $False
			}
		}
	}
	
	checkInteractiveOnly $packageName
	
}

function showToolTip ([System.Windows.Forms.Control] $hoverItem) {
	
	$tagDuration = 20000
	$tagText = $hoverItem.Tag
	[array] $nl = $tagText.Split('`r`n')
	[int] $height = $nl.Count * 2
	$tagLocation = New-Object System.Drawing.Point($hoverItem.Right,($hoverItem.Top - $height))
	$mainFormToolTip.Show($tagText, $mainForm, $tagLocation, $tagDuration)
	
}
	
#
# END FUNCTIONS
###################################################################################

###################################################################################
# BEGIN MAIN
#

# Copy the db file to local
copyDbToLocal

# Determine if the user is an admin
$isAdmin = $False
if ($groupMemberships.Contains($adminControlsGroup)) {
	$isAdmin = $True
}

# Get the configuration file parameters
readConfig

# Set up the packages array
[array] $packages = @()

# Set up the categories array
[array] $categories = getCategories

# Add the Windows Forms assembly to PowerShell
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

if (!$categories) {
	okBox "No categories yet. Please contact IT"
}

# Set up the main form properties
$mainForm = New-Object Windows.Forms.Form
$mainForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($icoFile)
$mainForm.Text = "PDQ AppStore"
$mainForm.Size = New-Object System.Drawing.Size(810,775)
$mainForm.StartPosition = "CenterScreen"
$mainForm.Opacity = 1
$mainForm.KeyPreview = $False
# Font styles are: Regular, Bold, Italic, Underline, Strikeout
$mainForm.Font = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)

# Set up the tooltip properties
$mainFormToolTip = New-Object System.Windows.Forms.ToolTip 
$mainFormToolTip.InitialDelay = 50000     
$mainFormToolTip.ReshowDelay = 20000 
$mainFormToolTip.IsBalloon = $False
$mainFormToolTip.ShowAlways = $False

# Category selection drop down
$mainFormCategoryDropDown = New-Object System.Windows.Forms.Combobox
$mainFormCategoryDropDown.Text = "Select a category"
$mainFormCategoryDropDown.DataSource = $categories
$mainFormCategoryDropDown.Size = New-Object System.Drawing.Size(165,30)
$mainFormCategoryDropDown.Location = New-Object System.Drawing.Point(10,22)
$mainFormCategoryDropDown.Add_SelectedValueChanged({
	$packages = getPackages $this.Text
	$mainForm.Refresh()
})

# Categories group box
$mainFormGroupBox1 = New-Object System.Windows.Forms.Groupbox
$mainFormGroupBox1.Text = "Category"
$mainFormGroupBox1.Size = New-Object System.Drawing.Size(185,60)
$mainFormGroupBox1.location = New-Object System.Drawing.Point(10,5)
$mainFormGroupBox1.controls.Add($mainFormCategoryDropDown)

$mainForm.Controls.Add($mainFormGroupBox1)

# Deploy now button
$mainFormButton1 = New-Object System.Windows.Forms.Button
$mainFormButton1.Text = "Deploy"
$mainFormButton1.Tag = "Clicking will queue the selected software for immediate deployment.`nIt can take up to 30 seconds for the server to start the process`nYou will be notified by pop-up message when deployment has completed."
$mainFormButton1.Size = New-Object System.Drawing.Size(75,30)
$mainFormButton1.Location = New-Object System.Drawing.Point(10,20)
$mainFormButton1.Font = New-Object System.Drawing.Font("Arial Narrow",14,[System.Drawing.FontStyle]::Regular)
$mainFormButton1.Add_Click({ 
	queueDeployment $mainFormPackageSelection.Text
	toggleDeploybuttons 0 $mainFormPackageSelection.Text
})
$mainFormButton1.Add_MouseHover({ showToolTip $this })
$mainFormButton1.Add_MouseLeave({ $mainFormToolTip.Hide($mainForm) })

# Deploy later button
$mainFormButton2 = New-Object System.Windows.Forms.Button
$mainFormButton2.Text = "Schedule"
$mainFormButton2.Tag = "Clicking will allow you to queue the selected software for later deployment.`nYou will be sent an e-mail when the deployment starts.`nNote: Your computer needs to be online at the time of deployment, but you do not need to be logged in."
$mainFormButton2.Size = New-Object System.Drawing.Size(85,30)
$mainFormButton2.Location = New-Object System.Drawing.Point(($mainFormButton1.Right + 5),20)
$mainFormButton2.Font = New-Object System.Drawing.Font("Arial Narrow",14,[System.Drawing.FontStyle]::Regular)
$mainFormButton2.Add_Click({ 
	schedulePicker $mainFormPackageSelection.Text
	toggleDeploybuttons 0 $mainFormPackageSelection.Text
})
$mainFormButton2.Add_MouseHover({ showToolTip $this })
$mainFormButton2.Add_MouseLeave({ $mainFormToolTip.Hide($mainForm) })

# View pending deployments
$mainFormButton3 = New-Object System.Windows.Forms.Button
$mainFormButton3.Text = "View Queue"
$mainFormButton3.Tag = "View pending deployments for this computer."
$mainFormButton3.Size = New-Object System.Drawing.Size(110,30)
$mainFormButton3.Location = New-Object System.Drawing.Point(($mainFormButton2.Right + 5),20)
$mainFormButton3.Font = New-Object System.Drawing.Font("Arial Narrow",14,[System.Drawing.FontStyle]::Regular)
$mainFormButton3.Add_Click({ viewDeployments })
$mainFormButton3.Add_MouseHover({ showToolTip $this })
$mainFormButton3.Add_MouseLeave({ $mainFormToolTip.Hide($mainForm) })

# View pending deployments
$mainFormButton4 = New-Object System.Windows.Forms.Button
$mainFormButton4.Text = "View Log"
$mainFormButton4.Tag = "View history of past deployments of the selected software for this computer."
$mainFormButton4.Size = New-Object System.Drawing.Size(90,30)
$mainFormButton4.Location = New-Object System.Drawing.Point(($mainFormButton3.Right + 5),20)
$mainFormButton4.Font = New-Object System.Drawing.Font("Arial Narrow",14,[System.Drawing.FontStyle]::Regular)
$mainFormButton4.Add_Click({ viewDeploymentHistory $mainFormPackageSelection.Text })
$mainFormButton4.Add_MouseHover({ showToolTip $this })
$mainFormButton4.Add_MouseLeave({ $mainFormToolTip.Hide($mainForm) })

# Actions group box
$mainFormGroupBox2 = New-Object System.Windows.Forms.Groupbox
$mainFormGroupBox2.text = "Actions"
$mainFormGroupBox2.Size = New-Object System.Drawing.Size(395,60)
$mainFormGroupBox2.location = New-Object System.Drawing.Point(($mainFormGroupBox1.Right + 5),5)
$mainFormGroupBox2.controls.AddRange(@($mainFormButton1,$mainFormButton2,$mainFormButton3,$mainFormButton4))

$mainForm.Controls.Add($mainFormGroupBox2)

# Package selection text box
$mainFormPackageSelection = New-Object System.Windows.Forms.ListBox
$mainFormPackageSelection.Size = New-Object System.Drawing.Size(755,330)
$mainFormPackageSelection.Location = New-Object System.Drawing.Point(10,20)
$mainFormPackageSelection.Add_SelectedValueChanged({
	$packages = getPackageDescription $this.Text
	toggleDeploybuttons 1 $this.Text
})
	
# Package selection group box
$mainFormGroupBox3 = New-Object System.Windows.Forms.Groupbox
$mainFormGroupBox3.text = "Software"
$mainFormGroupBox3.Size = New-Object System.Drawing.Size(775,360)
$mainFormGroupBox3.location = New-Object System.Drawing.Point(10,($mainFormGroupBox1.Bottom + 10))
$mainFormGroupBox3.controls.Add($mainFormPackageSelection)

$mainForm.Controls.Add($mainFormGroupBox3)

# Package description text box
$mainFormPackageDescription = New-Object System.Windows.Forms.TextBox
$mainFormPackageDescription.Size = New-Object System.Drawing.Size(755,255)
$mainFormPackageDescription.Location = New-Object System.Drawing.Point(10,20)
$mainFormPackageDescription.Multiline = $True
$mainFormPackageDescription.WordWrap = $True
$mainFormPackageDescription.AcceptsTab = $True
$mainFormPackageDescription.AcceptsReturn = $True
$mainFormPackageDescription.ScrollBars = 'Vertical'
if ($isAdmin -eq $False) {
	$mainFormPackageDescription.ReadOnly = $True
}

# Package description group box
$mainFormGroupBox4 = New-Object System.Windows.Forms.Groupbox
$mainFormGroupBox4.text = "Software Description"
$mainFormGroupBox4.Size = New-Object System.Drawing.Size(775,285)
$mainFormGroupBox4.location = New-Object System.Drawing.Point(10,($mainFormGroupBox3.Bottom + 10))
$mainFormGroupBox4.controls.Add($mainFormPackageDescription)

$mainForm.Controls.Add($mainFormGroupBox4)

if ($isAdmin -eq $True) {

	# Deploy now button
	$mainFormButton5 = New-Object System.Windows.Forms.Button
	$mainFormButton5.Text = "Update"
	$mainFormButton5.Tag = "Clicking will update the description for the selected software"
	$mainFormButton5.Size = New-Object System.Drawing.Size(80,30)
	$mainFormButton5.Location = New-Object System.Drawing.Point(10,20)
	$mainFormButton5.Font = New-Object System.Drawing.Font("Arial Narrow",14,[System.Drawing.FontStyle]::Regular)
	$mainFormButton5.Add_Click({ updateDescription $mainFormPackageSelection.Text $mainFormPackageDescription.Text })
	$mainFormButton5.Add_MouseHover({ showToolTip $this })
	$mainFormButton5.Add_MouseLeave({ $mainFormToolTip.Hide($mainForm) })

	# Deploy later button
	$mainFormButton6 = New-Object System.Windows.Forms.Button
	$mainFormButton6.Text = "Manage"
	$mainFormButton6.Tag = "Package management control panel"
	$mainFormButton6.Size = New-Object System.Drawing.Size(80,30)
	$mainFormButton6.Location = New-Object System.Drawing.Point(($mainFormButton5.Right + 5),20)
	$mainFormButton6.Font = New-Object System.Drawing.Font("Arial Narrow",14,[System.Drawing.FontStyle]::Regular)
	$mainFormButton6.Add_Click({ showAdminOptions })
	$mainFormButton6.Add_MouseHover({ showToolTip $this })
	$mainFormButton6.Add_MouseLeave({ $mainFormToolTip.Hide($mainForm) })

	
	# Admin group box
	$mainFormGroupBox5 = New-Object System.Windows.Forms.Groupbox
	$mainFormGroupBox5.text = "Admin"
	$mainFormGroupBox5.Size = New-Object System.Drawing.Size(185,60)
	$mainFormGroupBox5.location = New-Object System.Drawing.Point(($mainFormGroupBox2.Right + 5),5)
	$mainFormGroupBox5.controls.AddRange(@($mainFormButton5,$mainFormButton6))
	
	$mainForm.Controls.Add($mainFormGroupBox5)
	
}

if ($Cfg.Get_Item("AlwaysOnTop") -imatch "^y") {
	$mainForm.Topmost = $True
} else {
	$mainForm.Topmost = $False
}

[void] $mainForm.ShowDialog()

#
# END MAIN
###################################################################################
