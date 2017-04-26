'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2007
'
' NAME: KillWOW.vbs
'
' AUTHOR: Dave DeCoursey , Major Electric Supply, Inc.
' CREATE AT  : 5/7/2009
' MODIFY AT  : 26/04/2017 BY Lucas do Nascimento
'
' COMMENT: This will check a process and if it's using a certain amount of processor
'			percentage for over 1 minute, it will terminate the process
'
'==========================================================================
Option Explicit

CheckProcess

Function CheckProcess
	Const taskName = "splwow64"		' This is the name of the process to monitor
	Const taskToKill = "splwow64.exe"' We kill this one, notice 2 names depending on
									' on the system we're using
	Const percentUsed = 0			' This is the amount of processor it's allowed to use
	Const secondsToKill = 5			' This is how long it needs to be at or above percentUsed before killing
	Const sleepTime = 4000			' msec to sleep between Loop
						' 1000 * seconds (or minutes * 60 * 1000)
									
	Dim lastTimeUnder		' The last time the process was under the stated percentage
	Dim objWMIService		' The WMI Service object
	Dim objRefresher		' A SWbemRefresher object
	Dim colItems			' The process collection
	Dim objItem			' One process
	Dim idx				' Misc index
	Dim strComputer			' Computer to be checking
	Dim QueryWMI			' For another query
	Dim cProcesses			' A different list of processes
	Dim oProcess			' And each process in the above
	Dim sql				' Holds our Query string
	Dim oWMILocator			' what is says
	Dim oWMIService			' and again
	
	' Initialize the things we need
	strComputer = "."		' Local computer
	lastTimeUnder = Now		' Assume that we are starting below percentUsed
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strComputer & "\root\cimv2")
	Set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
	Set colItems = objRefresher.AddEnum _
		(objWMIService, "Win32_PerfFormattedData_PerfProc_Process").objectSet
	objRefresher.Refresh
	
	Do While True		' Loop until we kill the object
		For Each objItem In colItems
			If objItem.Name = taskName Then
				If objItem.PercentProcessorTime < percentUsed Then
					lastTimeUnder = Now
				Else
					If CallTimeSeconds(lastTimeUnder, Now) > secondsToKill Then
						sql = "SELECT * FROM WIN32_Process WHERE Name = '" & taskToKill & "'"
						set oWMILocator = CreateObject("WbemScripting.SWbemLocator")
						Set oWMIService = oWMILocator.ConnectServer(strComputer, "root/cimv2")
						Set cProcesses = oWMIService.ExecQuery(sql,,48)
						If Not IsObject(cProcesses) Then
							WScript.Echo " *** Could not retrieve process on " & strComputer
						Else
							For Each oProcess In CProcesses
								On Error Resume Next
								oProcess.Terminate()
								If Err <> 0 Then
									WScript.Echo " *** Error terminating on " & strComputer
									WScript.Echo " " & Err.Description
								End If
							Next
						End If
						Exit Function
					End If
				End If
			End if
		Next
		Dim objShell
		Set objShell = WScript.CreateObject( "WScript.Shell" )
		objShell.Exec("c:\Windows\splwow64.exe")
		Set objShell = Nothing
		WScript.Sleep sleepTime
		objRefresher.Refresh
	Loop
End function

Function CallTimeSeconds(startTime,endTime)
	Dim startHour
	Dim startMin
	Dim startSec
	Dim endHour
	Dim endMin
	Dim endSec
	Dim startingSeconds
	Dim endingSeconds
	
	startHour = Hour(startTime)
	startMin =  Minute(startTime)
	startSec = Second(startTime)
	endHour = Hour(endTime)
	endMin = Minute(endTime)
	endSec = Second(endTime)
	startingSeconds = (startSec + (startMin * 60) + ((startHour * 60)*60))
	endingSeconds = (EndSec + (endMin * 60) + ((endHour * 60)*60))
	CallTimeSeconds = endingSeconds - startingSeconds
End Function