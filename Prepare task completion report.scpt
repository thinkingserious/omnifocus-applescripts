-- This script is based on the code here: http://www.tuaw.com/2013/02/18/applescripting-omnifocus-send-completed-task-report-to-evernot/

-- Prepare a name for the new Evernote note
set theNoteName to "OmniFocus Completed Task Report"
set theNotebookName to ".Inbox"

-- Prompt the user to choose a scope for the report
activate
set theReportScope to choose from list {"Today", "Yesterday", "This Week", "Last Week", "This Month", "Last Month"} default items {"Yesterday"} with prompt "Generate a report for:" with title theNoteName
if theReportScope = false then return
set theReportScope to item 1 of theReportScope

-- Calculate the task start and end dates, based on the specified scope
set theStartDate to current date
set hours of theStartDate to 0
set minutes of theStartDate to 0
set seconds of theStartDate to 0
set theEndDate to theStartDate + (23 * hours) + (59 * minutes) + 59

if theReportScope = "Today" then
	set theDateRange to date string of theStartDate
else if theReportScope = "Yesterday" then
	set theStartDate to theStartDate - 1 * days
	set theEndDate to theEndDate - 1 * days
	set theDateRange to date string of theStartDate
else if theReportScope = "This Week" then
	repeat until (weekday of theStartDate) = Sunday
		set theStartDate to theStartDate - 1 * days
	end repeat
	repeat until (weekday of theEndDate) = Saturday
		set theEndDate to theEndDate + 1 * days
	end repeat
	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
else if theReportScope = "Last Week" then
	set theStartDate to theStartDate - 7 * days
	set theEndDate to theEndDate - 7 * days
	repeat until (weekday of theStartDate) = Sunday
		set theStartDate to theStartDate - 1 * days
	end repeat
	repeat until (weekday of theEndDate) = Saturday
		set theEndDate to theEndDate + 1 * days
	end repeat
	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
else if theReportScope = "This Month" then
	repeat until (day of theStartDate) = 1
		set theStartDate to theStartDate - 1 * days
	end repeat
	repeat until (month of theEndDate) is not equal to (month of theStartDate)
		set theEndDate to theEndDate + 1 * days
	end repeat
	set theEndDate to theEndDate - 1 * days
	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
else if theReportScope = "Last Month" then
	if (month of theStartDate) = January then
		set (year of theStartDate) to (year of theStartDate) - 1
		set (month of theStartDate) to December
	else
		set (month of theStartDate) to (month of theStartDate) - 1
	end if
	set month of theEndDate to month of theStartDate
	set year of theEndDate to year of theStartDate
	repeat until (day of theStartDate) = 1
		set theStartDate to theStartDate - 1 * days
	end repeat
	repeat until (month of theEndDate) is not equal to (month of theStartDate)
		set theEndDate to theEndDate + 1 * days
	end repeat
	set theEndDate to theEndDate - 1 * days
	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
end if

-- Begin preparing the task list as HTML
set theProgressDetail to "<html><body><h1>Completed Tasks</h1></br><b>" & theDateRange & "</b><br><hr><br>"
set theInboxProgressDetail to "<br>"

-- Retrieve a list of projects modified within the specified scope
set modifiedTasksDetected to false
tell application "OmniFocus"
	tell front document
		set theModifiedProjects to every flattened project where its modification date is greater than theStartDate
		-- Loop through any detected projects
		repeat with a from 1 to length of theModifiedProjects
			set theCurrentProject to item a of theModifiedProjects
			-- Retrieve any project tasks modified within the specified scope
			set theCompletedTasks to (every flattened task of theCurrentProject where its completed = true and completion date is greater than theStartDate and completion date is less than theEndDate and number of tasks = 0)
			-- Loop through any detected tasks
			if theCompletedTasks is not equal to {} then
				set modifiedTasksDetected to true
				-- Append the project name to the task list
				set theProgressDetail to theProgressDetail & "<h2>" & name of theCurrentProject & "</h2>" & return & "<br><ul>"
				repeat with b from 1 to length of theCompletedTasks
					set theCurrentTask to item b of theCompletedTasks
					-- Append the tasks's name to the task list
					set theProgressDetail to theProgressDetail & "<li>" & name of theCurrentTask & "</li>" & return
				end repeat
				set theProgressDetail to theProgressDetail & "</ul>" & return
			end if
		end repeat
		-- Include the OmniFocus inbox
		set theInboxCompletedTasks to (every inbox task where its completed = true and completion date is greater than theStartDate and completion date is less than theEndDate and number of tasks = 0)
		-- Loop through any detected tasks
		if theInboxCompletedTasks is not equal to {} then
			set modifiedTasksDetected to true
			-- Append the project name to the task list
			set theInboxProgressDetail to theInboxProgressDetail & "<h2>" & "Inbox" & "</h2>" & return & "<br><ul>"
			repeat with d from 1 to length of theInboxCompletedTasks
				-- Append the tasks's name to the task list
				set theInboxCurrentTask to item d of theInboxCompletedTasks
				set theInboxProgressDetail to theInboxProgressDetail & "<li>" & name of theInboxCurrentTask & "</li>" & return
			end repeat
			set theInboxProgressDetail to theInboxProgressDetail & "</ul>" & return
		end if

	end tell
end tell
set theProgressDetail to theProgressDetail & theInboxProgressDetail & "</body></html>"

-- Notify the user if no projects or tasks were found
if modifiedTasksDetected = false then
	display alert "OmniFocus Completed Task Report" message "No modified tasks were found for " & theReportScope & "."
	return
end if

-- Create the note in Evernote.
tell application "Evernote"
	activate
	set theReportDate to do shell script "date +%Y-%m-%d"
	set theNote to create note notebook theNotebookName title theReportDate & " :: " & theNoteName with html theProgressDetail
	open note window with theNote
end tell
