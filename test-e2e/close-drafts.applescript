-- Closes every Outlook compose window left over from a test run, answering the
-- "save this draft?" sheet with Discard Changes so nobody has to click it.
tell application "System Events" to tell process "Microsoft Outlook"
	set closedCount to 0
	repeat 20 times
		set target to missing value
		repeat with w in windows
			if (name of w) does not start with "Inbox" then
				set target to w
				exit repeat
			end if
		end repeat
		if target is missing value then exit repeat

		try
			click button 1 of target
		on error
			exit repeat
		end try
		delay 0.5

		-- Outlook asks about unsaved changes on a sheet.  Button names vary by
		-- version, so take whichever of these it offers.
		repeat with sheetName in {"Discard Changes", "Don't Save", "Delete", "No"}
			try
				click button (contents of sheetName) of sheet 1 of target
				delay 0.4
				exit repeat
			end try
		end repeat

		set closedCount to closedCount + 1
	end repeat
	return "closed " & closedCount
end tell
