set CurrDatetxt to short date string of date (short date string of (current date))
set dateYeartxt to year of (current date) as integer

if (month of (current date) as integer) < 10 then
	set dateMonthtxt to "0" & (month of (current date) as integer)
else
	set dateMonthtxt to month of (current date) as integer
end if

if (day of (current date) as integer) < 10 then
	set dateDaytxt to "0" & (day of (current date) as integer)
else
	set dateDaytxt to day of (current date) as integer
end if

set str_date to "" & dateYeartxt & "-" & dateMonthtxt & "-" & dateDaytxt

set theFilePath to "Macintosh HD:Users:Insomnia:Desktop:效率小结:小结" & str_date & ".md"

set due_Tasks to my OmniFocus_task_list()
my write_File(theFilePath, due_Tasks)

on OmniFocus_task_list()
	set CurrDate to date (short date string of (current date))
	set TomDate to date (short date string of ((current date) + days))
	set TomTomDate to date (short date string of ((current date) + days * 2))
	set Tom7Date to date (short date string of ((current date) + days * 7))
	set CurrDatetxt to short date string of date (short date string of (current date))
	set bigReturn to return & return
	set smallReturn to return
	set strText to "## 📮" & CurrDatetxt & " 效率小结：" & smallReturn
	tell application "OmniFocus"
		
		-- 1. 列出今日已完成的任务
		tell default document
			set refDueTaskList to a reference to (flattened tasks where (completion date > (CurrDate) and completion date ≤ TomDate and completed = true))
			set {lstName, lstProject, lstContext, lstDueDate} to {name, name of its containing project, name of its primary tag, due date} of refDueTaskList
			set strText to strText & bigReturn & "### 🏅 今日已完成 <span style=\"color:green\">" & length of lstName & " 项任务 </span>" & ":" & bigReturn
			set strText to strText & "|项目|任务|dateDue|" & smallReturn & "|--|--|--|" & smallReturn
			repeat with iTask from 1 to count of lstName
				set {strName, varProject, varContext, varDueDate} to {item iTask of lstName, item iTask of lstProject, item iTask of lstContext, item iTask of lstDueDate}
				if (varDueDate < (current date)) then
					set strDueDate to "**<span style=\"color:red\">" & short date string of varDueDate & "</span>**"
				else
					try
						set strDueDate to short date string of varDueDate
					on error
						set strDueDate to " " as string
					end try
				end if
				set strText to strText & "|" & "✅`" & varProject & "`" & "|" & "  " & strName & "  " & "|" & strDueDate & "|" & smallReturn
			end repeat
		end tell
		
		-- 2. 列出今日未完成的任务
		tell default document
			set refDueTaskList to a reference to (flattened tasks where due date ≤ TomDate and completed = false)
			set {lstName, lstProject, lstContext, lstDueDate} to {name, name of its containing project, name of its primary tag, due date} of refDueTaskList
			set strText to strText & bigReturn & "### ❗️ 今日未完成 <span style=\"color:orange\">" & length of lstName & " 项任务 </span>" & ":" & bigReturn
			set strText to strText & "|项目|任务|dateDue|" & smallReturn & "|--|--|--|" & smallReturn
			repeat with iTask from 1 to count of lstName
				set {strName, varProject, varContext, varDueDate} to {item iTask of lstName, item iTask of lstProject, item iTask of lstContext, item iTask of lstDueDate}
				if (varDueDate < (current date)) then
					set strDueDate to "**<span style=\"color:red\">" & short date string of varDueDate & "</span>**"
				else
					try
						set strDueDate to short date string of varDueDate
					on error
						set strDueDate to " " as string
					end try
				end if
				set strText to strText & "|" & "📌`" & varProject & "`" & "|" & "  " & strName & "  " & "|" & strDueDate & "|" & smallReturn
			end repeat
		end tell
		
		-- 3. 列出过期的任务
		tell default document
			set refDueTaskList to a reference to (flattened tasks where due date < CurrDate and completed = false)
			set {lstName, lstProject, lstContext, lstDueDate} to {name, name of its containing project, name of its primary tag, due date} of refDueTaskList
			set strText to strText & bigReturn & "### ‼️ 已有 <span style=\"color:red\">" & length of lstName & " 项任务严重延期 </span>" & ":" & bigReturn
			set strText to strText & "|项目|任务|dateDue|" & smallReturn & "|--|--|--|" & smallReturn
			repeat with iTask from 1 to count of lstName
				set {strName, varProject, varContext, varDueDate} to {item iTask of lstName, item iTask of lstProject, item iTask of lstContext, item iTask of lstDueDate}
				if (varDueDate < (current date)) then
					set strDueDate to "**<span style=\"color:red\">" & short date string of varDueDate & "</span>**"
				else
					try
						set strDueDate to short date string of varDueDate
					on error
						set strDueDate to " " as string
					end try
				end if
				set strText to strText & "|" & "📌`" & varProject & "`" & "|" & "  " & strName & "  " & "|" & strDueDate & "|" & smallReturn
			end repeat
		end tell
		
		-- 4. 列出明日已计划的任务
		tell default document
			set refDueTaskList to a reference to (flattened tasks where due date > TomDate and due date ≤ TomTomDate and completed = false)
			set {lstName, lstProject, lstContext, lstDueDate} to {name, name of its containing project, name of its primary tag, due date} of refDueTaskList
			set strText to strText & bigReturn & "### 📆 明日已计划 <span style=\"color:blue\">" & length of lstName & " 项任务 </span>" & ":" & bigReturn
			set strText to strText & "|项目|任务|dateDue|" & smallReturn & "|--|--|--|" & smallReturn
			repeat with iTask from 1 to count of lstName
				set {strName, varProject, varContext, varDueDate} to {item iTask of lstName, item iTask of lstProject, item iTask of lstContext, item iTask of lstDueDate}
				if (varDueDate < (current date)) then
					set strDueDate to "**<span style=\"color:red\">" & short date string of varDueDate & "</span>**"
				else
					try
						set strDueDate to short date string of varDueDate
					on error
						set strDueDate to " " as string
					end try
				end if
				set strText to strText & "|" & "💫`" & varProject & "`" & "|" & "  " & strName & "  " & "|" & strDueDate & "|" & smallReturn
			end repeat
		end tell
		
		-- 5. 列出未来七天的紧急和重要的任务
		tell default document
			set refDueTaskList to a reference to (flattened tasks where due date ≤ Tom7Date and completed = false)
			set {lstName, lstProject, lstContext, lstDueDate} to {name, name of its containing project, name of its primary tag, due date} of refDueTaskList
			set icount to 0
			repeat with iTask from 1 to count of lstName
				set varContext to item iTask of lstContext
				if varContext = "*Important" or varContext = "+Urgent" then
					set icount to icount + 1
				end if
			end repeat
			set strText to strText & bigReturn & "### 7️⃣ 未来七天共 <span style=\"color:purple\">" & icount & " 项**紧急/重要**任务 </span>" & ":" & bigReturn
			set strText to strText & "|项目|任务|dateDue|" & smallReturn & "|--|--|--|" & smallReturn
			repeat with iTask from 1 to count of lstName
				set {strName, varProject, varContext, varDueDate} to {item iTask of lstName, item iTask of lstProject, item iTask of lstContext, item iTask of lstDueDate}
				if varContext = "*Important" or varContext = "+Urgent" then
					if (varDueDate < (current date)) then
						set strDueDate to "**<span style=\"color:red\">" & short date string of varDueDate & "</span>**"
					else
						try
							set strDueDate to short date string of varDueDate
						on error
							set strDueDate to " " as string
						end try
					end if
					set strText to strText & "|" & "⚠️`" & varProject & "`" & "|" & "  " & strName & "  " & "|" & strDueDate & "|" & smallReturn
				end if
			end repeat
		end tell
		
		-- 6. 列出未指定日期且未完成的任务
		tell default document
			set refDueTaskList to a reference to (flattened tasks where due date = missing value and completed = false)
			set {lstName, lstProject, lstContext, lstDueDate} to {name, name of its containing project, name of its primary tag, due date} of refDueTaskList
			set strText to strText & bigReturn & "### 🈳 仍有 <span style=\"color:turquoise\">" & length of lstName & " 项任务未指定到期日 </span>" & ":" & bigReturn
			set strText to strText & "|项目|任务|dateDue|" & smallReturn & "|--|--|--|" & smallReturn
			repeat with iTask from 1 to count of lstName
				set {strName, varProject, varContext, varDueDate} to {item iTask of lstName, item iTask of lstProject, item iTask of lstContext, item iTask of lstDueDate}
				if (varDueDate < (current date)) then
					set strDueDate to "**<span style=\"color:red\">" & short date string of varDueDate & "</span>**"
				else
					try
						set strDueDate to short date string of varDueDate
					on error
						set strDueDate to " " as string
					end try
				end if
				set strText to strText & "|" & "❓`" & varProject & "`" & "|" & "  " & strName & "  " & "|" & strDueDate & "|" & smallReturn
			end repeat
		end tell
		
		
		
		
	end tell
	strText
end OmniFocus_task_list

--Export Task list to .MD file
on write_File(theFilePath, due_Tasks)
	set theText to due_Tasks
	set theFileReference to open for access theFilePath with write permission
	write theText to theFileReference as «class utf8»
	close access the theFileReference
end write_File


tell application "Typora"
	open file theFilePath
end tell
