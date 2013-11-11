-- modern.IE VM Installer v0.1.1
-- AppleScript to install modern.IE VMs with one click in VirtualBox for Mac.
--
-- Copyright © 2013 Jonathan Hogervorst. All rights reserved.
-- This code is licensed under MIT license. See LICENSE for details.

global scriptName
set scriptName to "modern.IE VM Installer"
set scriptVersionName to scriptName & " v0.1.1"

-- Names for all VMs
set VM_XP_IE6 to "XP - IE6"
set VM_XP_IE8 to "XP - IE8"
set VM_Vista_IE7 to "Vista - IE7"
set VM_Win7_IE8 to "Win7 - IE8"
set VM_Win7_IE9 to "Win7 - IE9"
set VM_Win7_IE10 to "Win7 - IE10"
set VM_Win7_IE11 to "Win7 - IE11 Preview"
set VM_Win8_IE10 to "Win8 - IE10"
set VM_Win81_IE11 to "Win8.1 - IE11 Preview"

set allVMs to {VM_XP_IE6, VM_XP_IE8, VM_Vista_IE7, VM_Win7_IE8, VM_Win7_IE9, VM_Win7_IE10, VM_Win7_IE11, VM_Win8_IE10, VM_Win81_IE11}

-- VM file types (OVA, OVA in ZIPs, OVA in SFXs)
set OVA_VMs to {VM_XP_IE8}
set ZIP_VMs to {VM_XP_IE6}
set SFX_VMs to {VM_Vista_IE7, VM_Win7_IE8, VM_Win7_IE9, VM_Win7_IE10, VM_Win7_IE11, VM_Win8_IE10, VM_Win81_IE11}

-- Based on: http://macscripter.net/viewtopic.php?id=18519
on trim(theString)
	set trimCharacters to {" ", tab, ASCII character 10, return, ASCII character 0}
	
	repeat until first character of theString is not in trimCharacters
		set theString to text 2 thru -1 of theString
	end repeat
	
	repeat until last character of theString is not in trimCharacters
		set theString to text 1 thru -2 of theString
	end repeat
	
	return theString
end trim

-- Source: http://vgable.com/blog/2009/04/24/how-to-check-if-an-application-is-running-with-applescript/
on ApplicationIsRunning(appName)
	tell application "System Events" to set appNameIsRunning to exists (processes where name is appName)
	return appNameIsRunning
end ApplicationIsRunning

-- Source: http://growl.info/documentation/applescript-support.php#growlisrunning
on ApplicationBundleIsRunning(bundleID)
	tell application "System Events" to set bundleIDIsRunning to (count of (every process whose bundle identifier is bundleID)) > 0
	return bundleIDIsRunning
end ApplicationBundleIsRunning

on activateVirtualBox()
	if not ApplicationIsRunning("VirtualBox") then
		-- Open the application if it's not running, and wait a few seconds
		tell application "VirtualBox" to activate
		delay 3
	end if
	
	-- Activate the application
	tell application "VirtualBox" to activate
end activateVirtualBox

-- Based on: http://growl.info/documentation/applescript-support.php#simpleNotificationSampleCode
on GrowlAvailable()
	if ApplicationBundleIsRunning("com.Growl.GrowlHelperApp") then
		tell application id "com.Growl.GrowlHelperApp"
			set the allNotificationsList to {"Install Progress Notification"}
			set the enabledNotificationsList to {"Install Progress Notification"}
			
			register as application scriptName all notifications allNotificationsList default notifications enabledNotificationsList
		end tell
		
		return true
	else
		return false
	end if
end GrowlAvailable

on GrowlNotify(theTitle, theDescription)
	if GrowlAvailable() then
		tell application id "com.Growl.GrowlHelperApp"
			close all notifications
			
			notify with name "Install Progress Notification" title theTitle description theDescription application name scriptName with sticky
		end tell
	end if
end GrowlNotify

-- Based on: http://stackoverflow.com/a/8298899/251760
on filename(thePath)
	set AppleScript's text item delimiters to "/"
	return the last item of (thePath's text items)
end filename

-- Ask user for VM
set selectedVM to (choose from list allVMs OK button name "OK" cancel button name "Cancel" with title scriptVersionName with prompt "Which IE VM do you want to install?")

if allVMs contains selectedVM then
	set selectedVM_URLs to {}
	
	-- Determine file URLs based on chosen VM
	if (selectedVM = {VM_XP_IE6}) then
		set selectedVM_URLs to {"https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE6_WinXP.ova.zip"}
	else if (selectedVM = {VM_XP_IE8}) then
		set selectedVM_URLs to {"https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE8_XP/IE8.XP.For.MacVirtualBox.ova"}
	else if (selectedVM = {VM_Vista_IE7}) then
		set selectedVM_URLs to {"https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE7_Vista/IE7.Vista.For.MacVirtualBox.part1.sfx", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE7_Vista/IE7.Vista.For.MacVirtualBox.part2.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE7_Vista/IE7.Vista.For.MacVirtualBox.part3.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE7_Vista/IE7.Vista.For.MacVirtualBox.part4.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE7_Vista/IE7.Vista.For.MacVirtualBox.part5.rar"}
	else if (selectedVM = {VM_Win7_IE8}) then
		set selectedVM_URLs to {"https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE8_Win7/IE8.Win7.For.MacVirtualBox.part1.sfx", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE8_Win7/IE8.Win7.For.MacVirtualBox.p MacVirtualBox.part2.rar art2.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE8_Win7/IE8.Win7.For.MacVirtualBox.part3.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE8_Win7/IE8.Win7.For.MacVirtualBox.part4.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE8_Win7/IE8.Win7.For.MacVirtualBox.part5.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE8_Win7/IE8.Win7.For.MacVirtualBox.part6.rar"}
	else if (selectedVM = {VM_Win7_IE9}) then
		set selectedVM_URLs to {"https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE9_Win7/IE9.Win7.For.MacVirtualBox.part1.sfx", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE9_Win7/IE9.Win7.For.MacVirtualBox.part2.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE9_Win7/IE9.Win7.For.MacVirtualBox.part3.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE9_Win7/IE9.Win7.For.MacVirtualBox.part4.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE9_Win7/IE9.Win7.For.MacVirtualBox.part5.rar"}
	else if (selectedVM = {VM_Win7_IE10}) then
		set selectedVM_URLs to {"https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE10_Win7/IE10.Win7.For.MacVirtualBox.part1.sfx", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE10_Win7/IE10.Win7.For.MacVirtualBox.part2.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE10_Win7/IE10.Win7.For.MacVirtualBox.part3.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE10_Win7/IE10.Win7.For.MacVirtualBox.part4.rar"}
	else if (selectedVM = {VM_Win7_IE11}) then
		set selectedVM_URLs to {"https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE11_Win7/IE11.Win7.For.MacVirtualBox.part1.sfx", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE11_Win7/IE11.Win7.For.MacVirtualBox.part2.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE11_Win7/IE11.Win7.For.MacVirtualBox.part3.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE11_Win7/IE11.Win7.For.MacVirtualBox.part4.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE11_Win7/IE11.Win7.For.MacVirtualBox.part5.rar"}
	else if (selectedVM = {VM_Win8_IE10}) then
		set selectedVM_URLs to {"https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE10_Win8/IE10.Win8.For.MacVirtualBox.part1.sfx", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE10_Win8/IE10.Win8.For.MacVirtualBox.part2.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE10_Win8/IE10.Win8.For.MacVirtualBox.part3.rar"}
	else if (selectedVM = {VM_Win81_IE11}) then
		set selectedVM_URLs to {"https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE11_Win81/IE11.Win8.1Preview.For.MacVirtualBox.part1.sfx", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE11_Win81/IE11.Win8.1Preview.For.MacVirtualBox.part2.rar", "https://az412801.vo.msecnd.net/vhd/IEKitV1_Final/VirtualBox/OSX/IE11_Win81/IE11.Win8.1Preview.For.MacVirtualBox.part3.rar"}
	end if
	
	set tempPath to POSIX path of (path to temporary items)
	set cancelled to 0
	set success to 0
	set errorText to "Unknown error"
	set errorDescription to ""
	
	-- Download all files
	repeat with i from 1 to count of selectedVM_URLs
		set selectedVM_URL to item i of selectedVM_URLs
		set selectedVM_URL_filename to filename(selectedVM_URL)
		
		-- Remove file if existing
		do shell script "rm -f '" & tempPath & selectedVM_URL_filename & "'"
		
		-- Get file size and start cURL download in background
		set selectedVM_URL_size to word 3 of (do shell script "curl -s -I '" & selectedVM_URL & "' | grep 'Content-Length'") as number
		set selectedVM_URL_CURL_PID to (do shell script "curl '" & selectedVM_URL & "' -o '" & tempPath & selectedVM_URL_filename & "' -s > /dev/null 2>&1 & echo $!")
		
		set selectedVM_URL_sizeMB to round (selectedVM_URL_size / 1000 / 1000)
		
		-- While the download is active
		repeat until selectedVM_URL_CURL_PID = ""
			try
				-- Check whether the cURL process is still running
				do shell script "kill -s 0 " & selectedVM_URL_CURL_PID
			on error
				-- An error is returned when the process does not exist (i.e., when cURL finished)
				set selectedVM_URL_CURL_PID to ""
			end try
			
			-- If the cURL process is still running and Growl is available
			if selectedVM_URL_CURL_PID ≠ "" then
				try
					-- Get size of download so far
					set selectedVM_URL_currentSize to word 1 of (do shell script "du -s -k '" & tempPath & selectedVM_URL_filename & "'") as number
				on error
					set selectedVM_URL_currentSize to 0
				end try
				
				set selectedVM_URL_currentSize to selectedVM_URL_currentSize * 1024
				set selectedVM_URL_currentSizeMB to round (selectedVM_URL_currentSize / 1000 / 1000)
				set selectedVM_URL_currentSizePercentage to round (selectedVM_URL_currentSize / selectedVM_URL_size * 100)
				
				-- Display progress
				GrowlNotify("Downloading file " & i & " of " & (count of selectedVM_URLs) & "…", "Downloaded " & selectedVM_URL_currentSizeMB & "MB of " & selectedVM_URL_sizeMB & "MB (" & selectedVM_URL_currentSizePercentage & "%)")
			end if
			
			delay 1
		end repeat
		
		-- If the user cancelled, all remaining files will be skipped
		if cancelled = 1 then
			exit repeat
		end if
	end repeat
	
	-- Is the download completed?
	if cancelled = 0 then
		set selectedVM_OVA_path to ""
		set selectedVM_OVA_pathFound to ""
		
		set selectedVM_URL to item 1 of selectedVM_URLs
		set selectedVM_URL_filename to filename(selectedVM_URL)
		
		-- Get the (path to the) OVA file
		if OVA_VMs contains selectedVM then
			set selectedVM_OVA_path to tempPath & selectedVM_URL_filename
		else if ZIP_VMs contains selectedVM then
			-- Extract the ZIP
			GrowlNotify("Extracting ZIP file…", "")
			set selectedVM_OVA_path to trim(do shell script "unzip -o '" & tempPath & selectedVM_URL_filename & "' -d '" & tempPath & "' | grep 'inflating:' | sed 's/inflating://g'")
		else if SFX_VMs contains selectedVM then
			-- 'Extract' (execute) the SFX
			GrowlNotify("Extracting SFX file…", "")
			do shell script "chmod +x '" & tempPath & selectedVM_URL_filename & "'"
			set selectedVM_OVA_path to tempPath & trim(do shell script "cd '" & tempPath & "'; './" & selectedVM_URL_filename & "' | grep 'OK' | egrep 'Extracting|\\.\\.\\.' | sed 's/^\\.\\.\\.//g' | sed 's/^Extracting//g' | sed 's/\\.ova.*$/.ova/g'")
		end if
		
		-- Check whether the OVA file really exists
		try
			set selectedVM_OVA_pathFound to do shell script "test -f '" & selectedVM_OVA_path & "' && echo 'Found'"
		end try
		
		if selectedVM_OVA_pathFound = "Found" then
			set selectedVM_OVA_installResult to ""
			
			try
				-- Import the file in VirtualBox
				GrowlNotify("Importing OVA file into VirtualBox…", "")
				set selectedVM_OVA_installResult to do shell script "VBoxManage import '" & selectedVM_OVA_path & "'"
				
				-- Check whether the installation was successful
				if "Successfully imported the appliance." is in selectedVM_OVA_installResult then
					set success to 1
				else
					-- Import somehow failed (without an error)
					set errorText to "VM file could not be imported in VirtualBox." & return & return & "VirtualBox output:"
					set errorDescription to selectedVM_OVA_installResult
				end if
			on error selectedVM_OVA_installError number selectedVM_OVA_installErrorNumber
				-- Is the VirtualBox command unknown (i.e., are the VirtualBox command line utilities not installed)?
				if selectedVM_OVA_installErrorNumber = 127 then
					-- Open/activate VirtualBox and open the OVA file
					activateVirtualBox()
					tell application "VirtualBox" to open selectedVM_OVA_path
					set success to 1
				else
					-- VirtualBox returned an error; the import failed
					set errorText to "VM file could not be imported in VirtualBox." & return & return & "VirtualBox error:"
					set errorDescription to selectedVM_OVA_installError
				end if
			end try
		else
			set errorText to "OVA file not found"
		end if
		
		-- Was the import successful?
		if success = 1 then
			GrowlNotify("Installation completed!", "")
			activateVirtualBox()
		else
			-- The import was not successful; show the error message
			try
				if errorDescription = "" then
					display dialog errorText buttons {"OK"} cancel button "OK" with title scriptVersionName
				else
					display dialog errorText default answer errorDescription buttons {"OK"} cancel button "OK" with title scriptVersionName
				end if
			end try
		end if
		
		-- Remove (extracted) OVA file
		do shell script "rm -f '" & selectedVM_OVA_path & "'"
	end if
	
	-- Remove downloaded files
	repeat with i from 1 to count of selectedVM_URLs
		set selectedVM_URL to item i of selectedVM_URLs
		set selectedVM_URL_filename to the last item of (selectedVM_URL's text items)
		do shell script "rm -f '" & tempPath & selectedVM_URL_filename & "'"
	end repeat
end if