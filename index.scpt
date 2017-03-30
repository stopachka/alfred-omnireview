on run argv
  -- set the report scope
  if count of argv is 0
    set theReportScope to "this-week"
  else
    set theReportScope to item 1 of argv
  end if

  -- Calculate the task start and end dates, based on the specified scope
  set theStartDate to current date
  set hours of theStartDate to 0
  set minutes of theStartDate to 0
  set seconds of theStartDate to 0
  set theEndDate to theStartDate + (23 * hours) + (59 * minutes) + 59

  if theReportScope = "today" then
  	set theDateRange to date string of theStartDate
  else if theReportScope = "yesterday" then
  	set theStartDate to theStartDate - 1 * days
  	set theEndDate to theEndDate - 1 * days
  	set theDateRange to date string of theStartDate
  else if theReportScope = "this-week" then
  	repeat until (weekday of theStartDate) = Sunday
  		set theStartDate to theStartDate - 1 * days
  	end repeat
  	repeat until (weekday of theEndDate) = Saturday
  		set theEndDate to theEndDate + 1 * days
  	end repeat
  	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
  else if theReportScope = "last-week" then
  	set theStartDate to theStartDate - 7 * days
  	set theEndDate to theEndDate - 7 * days
  	repeat until (weekday of theStartDate) = Sunday
  		set theStartDate to theStartDate - 1 * days
  	end repeat
  	repeat until (weekday of theEndDate) = Saturday
  		set theEndDate to theEndDate + 1 * days
  	end repeat
  	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
  else if theReportScope = "this-month" then
  	repeat until (day of theStartDate) = 1
  		set theStartDate to theStartDate - 1 * days
  	end repeat
  	repeat until (month of theEndDate) is not equal to (month of theStartDate)
  		set theEndDate to theEndDate + 1 * days
  	end repeat
  	set theEndDate to theEndDate - 1 * days
  	set theDateRange to (date string of theStartDate) & " through " & (date string of theEndDate)
  end if

  -- Prepare a name for the new note
  set theNoteName to "OmniReview: " & (short date string of theStartDate) & " - " & (short date string of theEndDate)

  -- Begin preparing the task list as HTML.
  set theProgressDetail to "<html><body><h1>Completed Tasks</h1><br><b>" & theDateRange & "</b><br><hr><br>"

  -- Retrieve a list of projects modified within the specified scope
  set modifiedTasksDetected to false
  tell application "OmniFocus"
  	tell front document
  		set theModifiedProjects to every flattened project where its modification date is greater than theStartDate and modification date is less than theEndDate

  		-- Loop through any detected projects
  		repeat with a from 1 to length of theModifiedProjects
  			set theCurrentProject to item a of theModifiedProjects

  			-- Retrieve any project tasks modified within the specified scope
  			set theCompletedTasks to (every flattened task of theCurrentProject where its completed = true and modification date is greater than theStartDate and modification date is less than theEndDate and number of tasks = 0)

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
  	end tell
  end tell
  set theProgressDetail to theProgressDetail & "</body></html>"

  -- Notify the user if no projects or tasks were found
  if modifiedTasksDetected = false then
  	display alert "OmniFocus Completed Task Report" message "No modified tasks were found for " & theReportScope & "."
  	return
  end if

  -- Create the note in Evernote.
  tell application "Evernote"
  	activate
  	set theNote to create note notebook "Vault" title theNoteName with html theProgressDetail
  	open note window with theNote
  end tell
end run
