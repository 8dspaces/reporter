'''<summary>This file holds the base classes for an engine and a filter</summary>
''' <datecreated>01/09/2007</datecreated>
'''<version>0.9</version>


Option Explicit
'ReporterManager Project home at  www.AdvancedQTP.com

'A template class for a Filter
''' <summary>
''' A Base class for a filter
''' </summary>
''' <author>Yaron Assa</author>
''' <datecreated>01/09/2007</datecreated>
''' <remarks>This class is essential for the intellisense of the whole project</remarks>
Class XFilter	'Change X to the name of the class (keep the "filter" part)
'Written by Yaron Assa, 1/8/07. Project home at www.AdvancedQTP.com

	'I find that keeping the settings withing a dictionary object is the best way to manage things
	'Feel free to change this mechanism, it's invisible to the outside world
	''' <summary>
	''' Holds a link to the reporter manager
	''' </summary>
	''' <value type="ReporterManagerClass"/>
	Public Manager 'Holds a link to the reporter manager
	''' <summary>
	''' Holds a link to the relevant engine
	''' </summary>
	''' <value type="XEngine"/>
	Public Engine 'Holds a link to the relevant engine
	
	''' <value type="Scripting.Dictionary"/>
	''' <summary>Will hold the settings dictionary</summary>	
	Private oSettings	'Will hold the settings dictionary
	
	''' <summary>
	''' ID for the type of the filter
	''' </summary>
	''' <value type="String"/>
	''' <example>"DB"</example>	
	Public Property Get FilterType
		FilterType = "Name" 'Change Name to the name of the class		
	End Property
	
	''' <summary>
	''' The Engine's Version
	''' </summary>
	''' <value type="String"/>
	''' <example>"0.8"</example>	
	Public Property Get Version
		Version = "9.0" 'Change to the filter version. Must be 9+
	End Property
	
	''' <summary>
	''' The minimum version of the manager, required for this filter to work
	''' </summary>
	''' <value type="String"/>
	''' <example>"0.8"</example>	
	Public Property Get MinimunManagerVersion
		MinimunManagerVersion = "0.0" 'The minimum version of the manager, required for this filter to work
	End Property

	'*************************************************************'
	'	Function StartFilter    (Public)
	'	Goal 								Makes the filter ready for use
	'	In Params					   sSettings - The Filters init settings as "Field1>Value1|Field2>Value2|...
	'	Out Params					Boolean - True/False
	'*************************************************************'  
	''' <summary>
	''' Makes the filter ready for use
	''' </summary>
	''' <param name="sSettings" type="String">The Filter's init settings as "Field1>Value1|Field2>Value2|...</param>
	''' <returns type="Boolean">True if succeeded to start the filter, False if failed</returns>
	Public Function StartFilter(sSettings)
	   ''' <value type="Boolean"/>
	   Dim bResult
	   
	   Load_Settings(sSettings)
	   'Add more filter init code here	   

	   bResult = AllowCreate 'Can we start the filter?
	   
	   StartFilter = bResult
	End Function

	'*************************************************************'
	'	Function Check    (Public)
	'	Goal 								Checks if an event matches the filter's criteria
   '	In Params                      sStatus - The status of the event : "Fail", "Pass", "Warning" or custom
   '											  sStepName - The name of the event (e.g. "Validate Field Value")
   '											  sExpected - The expected result (e.g. "Value = Something")
   '											  sActual - The actual result (e.g. "Value = Something Else")
   '											  sDetauls - Possile details about the event
	'	Out Params					Boolean - Does it match the criteria?
	'*************************************************************'  
''' <summary>Checks if an event matches the filter's criteria</summary>
''' <param name="sStatus" type="String" paramlist="StatusType">The event's status</param>
''' <paramlist name="StatusType">
''' <item>"Fail"<description>The Step Failed</description></item>
''' <item>"Pass"<description>The Step Passed</description></item>
''' <item>"Warning"<description>A Warning</description></item>
''' <item>"General"<description>Just report the event</description></item>
''' </paramlist>
''' <param name="sStepName" type="String">The event's step name</param>
''' <param name="sExpected" type="String">The event's expected result</param>
''' <param name="sActual" type="String">The event's actual result</param>
''' <param name="sDetails" type="String">The event's details</param>
''' <returns type="Boolean">Returns True if the event matches, otherwise False</returns>
	Public Function Check(sStatus, sStepName,sExpected,sActual, sDetails)
	   ''' <value type="Boolean"/>
	   Dim bResult

	   bResult = False
		'Implement Check code here!
		'This is the heart of the filter class

	   Check = bResult
	End Function

	'*************************************************************'
	'	Sub Init_Settings    (Private)
	'	Goal 								Specifies default values of the settings object
	'	In Params						None
	'	Out Params					None
	'*************************************************************'  
	''' <summary>
	''' Specifies default values of the settings object
	''' </summary>
	Private Sub  Init_Settings
	   Set oSettings = CreateObject("Scripting.Dictionary")
		'Add Default Settings Here
		'e.g. oSettings.Add "ClearExisting" , "True"
		'Make sure all mandatory fields have default values!
	End Sub

   '*************************************************************'
   '	Function AllowCreate  		  (Private)
   '	Goal 								Checks if there would be a problem to create the reporter engine
   '	In Params						None
   '	Out Params					Boolean - Yes/No (Yes = OK to create)
   '*************************************************************'
	''' <summary>
	''' Checks if there would be a problem to create the reporter engine
	''' </summary>
	''' <returns type="Boolean">True if it's OK to create the filter, False otherwise</returns>
	Private Function AllowCreate
	   ''' <value type="Boolean"/>
	   Dim bResult

	   bResult = CheckVersion
	   'Add more pass/Fail conditions here
		
	   AllowCreate = bResult
	End Function

   '*************************************************************'
   '	Function CheckVersion  		  (Private)
   '	Goal 								Checks if the manager and engine versions are ok
   '	In Params						None
   '	Out Params					Boolean - Yes/No (Yes = OK to create)
   '*************************************************************'
	''' <summary>
	''' Checks if the manager and engine versions are ok
	''' </summary>
	''' <returns type="Boolean">True if the filter passes the version check, False otherwise</returns>
	Function CheckVersion
	   ''' <value type="Integer"/>
	   Dim iDiff
	   ''' <value type="String"/>
	   Dim sEngineVersion
	   ''' <value type="String"/>
	   Dim sManagerVersion

	   'Check the manager version is good enough
	   sEngineVersion = Replace(CStr(Me.MinimunManagerVersion), ".", "")
	   sManagerVersion = Replace(CStr(Me.Manager.Version), ".","")

	   iDiff = Len(sEngineVersion) - Len(sManagerVersion)

	   If iDiff > 0 Then
			sManagerVersion = sManagerVersion & String(iDiff, "0")
		Else
			sEngineVersion = sEngineVersion & String(iDiff, "0")	
	   End If

	   If CLng(sEngineVersion) > CLng(sManagerVersion)  Then
		   Reporter.ReportEvent MicFail, "Start Engine", "Failed Version Check" & vbCrLf & "Engine Expected " & Me.MinimunManagerVersion & _
																										vbCrLf & "But manager was actually " & Me.Manager.Version
		   CheckVersion = False
		   Exit Function
	   End If

	   'Check the engine version is good enough
	   sEngineVersion = Replace(CStr(Me.Version), ".", "")
	   sManagerVersion = Replace(CStr(Me.Manager.MinimumEngineVersion), ".","")

	   iDiff = Len(sEngineVersion) - Len(sManagerVersion)

	   If iDiff > 0 Then
			sManagerVersion = sManagerVersion & String(iDiff, "0")
		Else
			sEngineVersion = sEngineVersion & String(iDiff, "0")	
	   End If

	   If CLng(sEngineVersion) < CLng(sManagerVersion)  Then
		   Reporter.ReportEvent MicFail, "Start Engine", "Failed Version Check" & vbCrLf & "Manager expected " & Me.Manager.MinimumEngineVersion &_
																									  vbCrLf & "But engine actually was " & Me.Version
		   CheckVersion = False
		   Exit Function
	   End If

		CheckVersion = True
	End Function

	'*************************************************************'
	'	Sub Class_Initialize    (Private)
	'	Goal 								Setup Initialize event.
	'	In Params						None
	'	Out Params					None
	'*************************************************************'  
''' <summary>
''' Setup Initialize event
''' </summary>
   Private Sub Class_Initialize()   		
		Init_Settings
		'Add Initialization events here
   End Sub

	'*************************************************************'
	'	Sub Class_Terminate    (Private)
	'	Goal 								Setup Terminate event.
	'	In Params						None
	'	Out Params					None
	'*************************************************************'  
''' <summary>
''' Setup Terminate event
''' </summary>
	Private Sub Class_Terminate 
		Set oSettings = Nothing
		'Add Termination  events here		
   End Sub

	'*************************************************************'
	'	Sub Load_Settings    (Private)
	'	Goal 								Loads sSettings into the oSettings object
	'	In Params						sSettings - The Filters init settings as "Field1>Value1|Field2>Value2|...
	'	Out Params					None
	'*************************************************************'  
	''' <summary>
	''' Loads sSettings into the oSettings object
	''' </summary>
	''' <param name="sSettings" type="String">The Filter's init settings as "Field1>Value1|Field2>Value2|...</param>
	Private Sub Load_Settings(sSettings) 
	   '***************************' No Need to change this sub! ***************************'
		''' <value type="Array"/>
		Dim ReplaceArray
		''' <value type="Array"/>
		Dim arrData
		''' <value type="Array"/>
		Dim arrSettingData
		''' <value type="Integer"/>
		Dim i

		'Protect >> and || chars
		ReplaceArray = Array ("ABCDEFGHIJKLMNOP1", "ABCDEFGHIJKLMNOP2")

			sSettings  = Replace(sSettings, "||", ReplaceArray(0))
			sSettings  = Replace(sSettings, ">>", ReplaceArray(1))
			
		   arrData = Split(sSettings, "|") 'Split to individual fields
	
			   For i = 0 To UBound(arrData) 'Loop through fields
				   If arrData(i) <> "" Then
						arrSettingData = Split(arrData(i), ">") 'Break to Field and to Value
						
						'Protect | char
						arrSettingData(0) = Replace(arrSettingData(0), ReplaceArray(0), "|")
						arrSettingData(1) = Replace(arrSettingData(1), ReplaceArray(0), "|")

						'Protect > char
						arrSettingData(0) = Replace(arrSettingData(0), ReplaceArray(1), ">")
						arrSettingData(1) = Replace(arrSettingData(1), ReplaceArray(1), ">")

						If oSettings.Exists(arrSettingData(0)) Then 'There's a default value for this field
							oSettings(arrSettingData(0)) = arrSettingData(1) 'Change it to the user's value
						Else
							oSettings.Add arrSettingData(0), arrSettingData(1) 'Create new setting
						End If
						
					End If
					
			   Next
	
	End Sub

	''' <summary>
	''' Forces the filter to close and clear all its references
	''' </summary>
	Public Sub ClearFilter
		Class_Terminate
	End Sub
End Class

'A template class for a reporter engine
''' <summary>
''' A Base class for an Engine
''' </summary>
''' <author>Yaron Assa</author>
''' <datecreated>01/09/2007</datecreated>
''' <remarks>This class is essential for the intellisense of the whole project</remarks>
Class XEngine 'Change X to the name of the class (keep the "Engine" part)
'Written by Yaron Assa, 1/8/07. Project home at www.AdvancedQTP.com

   	'I find that keeping the settings withing a dictionary object is the best way to manage things
	'Feel free to change this mechanism, it's invisible to the outside world
	''' <summary>
	''' Holds a link to the reporter manager
	''' </summary>
	''' <value type="ReporterManagerClass"/>
	Public Manager 'Holds a link to the reporter manager
	''' <summary>
	''' The user logical name for an instance of this engine (in run-time)
	''' </summary>
	''' <value type="String"/>
	Public Name

	''' <value type="Scripting.Dictionary"/>
	''' <summary>Will hold the settings dictionary</summary>	
	Private oSettings
	''' <value type="Object"/>
	''' <summary>Holds the engine's main object</summary>	
	''' <remarks>You can use it or not, according to your specific implementation</remarks>	
	Private oEngine	'Holds the engine's main object

	''' <summary>
	''' ID for the type of the engine
	''' </summary>
	''' <value type="String"/>
	''' <example>"DB"</example>	
	Public Property Get EngineType
		EngineType = "Name" 'Add Engine Type Here
	End Property

''' <summary>
	''' The Engine's Version
	''' </summary>
	''' <value type="String"/>
	''' <example>"0.8"</example>	
	Public Property Get Version
		Version = "9.0" 'Change to the engine version. Must be 9+
	End Property

''' <summary>
	''' The minimum version of the manager, required for this engine to work
	''' </summary>
	''' <value type="String"/>
	''' <example>"0.8"</example>	
	Public Property Get MinimunManagerVersion
		MinimunManagerVersion = "0.0" 'The minimum version of the manager, required for this engine to work
	End Property

	'*************************************************************'
	'	Function StartEngine    (Public)
	'	Goal 								Makes the engine ready for use
	'	In Params					   sSettings - The Engine's init settings as "Field1>Value1|Field2>Value2|...
	'	Out Params					Boolean - Didn't the engine manage to start?
	'*************************************************************'  
	''' <summary>
	''' Makes the engine ready for use
	''' </summary>
	''' <param name="sSettings" type="String">The Engine's init settings as "Field1>Value1|Field2>Value2|...</param>
	''' <returns type="Boolean">True if succeeded to start the engine, False if failed</returns>
	Public Function StartEngine(sSettings)
	   ''' <value type="Boolean"/>
	   Dim bResult

		Load_Settings(sSettings)
		'More engine start code here

	   bResult = AllowCreate 'Can we create this engine?
		
	   StartEngine = bResult
	End Function

	'*************************************************************'
	'	Function StopEngine    (Public)
	'	Goal 								Stops the engine
	'	In Params					   None
	'	Out Params					None
	'*************************************************************'  
	''' <summary>
	''' Stops the engine
	''' </summary>
	Public Sub StopEngine
			Set oEngine = Nothing					
			Set oSettings = Nothing
			'More engine stop code here
	End Sub

   '*************************************************************'
   '	Sub Report  (Public)
   '	Goal 								Reports an event 
   '	In Params					   sStatus - The status of the event : "Fail", "Pass", "Warning" or custom
   '											  sStepName - The name of the event (e.g. "Validate Field Value")
   '											  sExpected - The expected result (e.g. "Value = Something")
   '											  sActual - The actual result (e.g. "Value = Something Else")
   '											  sDetauls - Possile details about the event
   '	Out Params					None
   '*************************************************************'   	    
''' <summary>
''' Reports an event to all possible reporters
''' </summary>
''' <param name="sStatus" type="String" paramlist="StatusType">The event's status</param>
''' <paramlist name="StatusType">
''' <item>"Fail"<description>The Step Failed</description></item>
''' <item>"Pass"<description>The Step Passed</description></item>
''' <item>"Warning"<description>A Warning</description></item>
''' <item>"General"<description>Just report the event</description></item>
''' </paramlist>
''' <param name="sStepName" type="String">The event's step name</param>
''' <param name="sExpected" type="String">The event's expected result</param>
''' <param name="sActual" type="String">The event's actual result</param>
''' <param name="sDetails" type="String">The event's details</param>
	Public Sub Report (sStatus, sStepName,sExpected,sActual, sDetails)
	   'Enter Report Code Here
	   'This is the heart of the engine's class!
	End Sub

	'*************************************************************'
	'	Sub Init_Settings    (Private)
	'	Goal 								Specifies default values of the settings object
	'	In Params						None
	'	Out Params					None
	'*************************************************************'  
	''' <summary>
	''' Specifies default values of the settings object
	''' </summary>
	Private Sub  Init_Settings
	   Set oSettings = CreateObject("Scripting.Dictionary")
		'Add Default Settings Here
		'e.g. oSettings.Add "File" , "C:\QTPLog.txt"
		'Make sure all mandatory fields have default values!
	End Sub

   '*************************************************************'
   '	Function AllowCreate  		  (Private)
   '	Goal 								Checks if there would be a problem to create the reporter engine
   '	In Params						None
   '	Out Params					Boolean - Yes/No (Yes = OK to create)
   '*************************************************************'
	''' <summary>
	''' Checks if there would be a problem to create the reporter engine
	''' </summary>
	''' <returns type="Boolean">True if it's OK to create the filter, False otherwise</returns>
	Private Function AllowCreate
	   ''' <value type="Boolean"/>
	   Dim bResult

	   bResult = CheckVersion
	   'Add more pass/fail conditions here
		
	   AllowCreate = bResult
	End Function

   '*************************************************************'
   '	Function CheckVersion  		  (Private)
   '	Goal 								Checks if the manager and engine versions are ok
   '	In Params						None
   '	Out Params					Boolean - Yes/No (Yes = OK to create)
   '*************************************************************'
	''' <summary>
	''' Checks if the manager and engine versions are ok
	''' </summary>
	''' <returns type="Boolean">True if the engine passes the version check, False otherwise</returns>
	Function CheckVersion
	   ''' <value type="Integer"/>
	   Dim iDiff
	   ''' <value type="String"/>
	   Dim sEngineVersion
	   ''' <value type="String"/>
	   Dim sManagerVersion

	   'Check the manager version is good enough
	   sEngineVersion = Replace(CStr(Me.MinimunManagerVersion), ".", "")
	   sManagerVersion = Replace(CStr(Me.Manager.Version), ".","")

	   iDiff = Len(sEngineVersion) - Len(sManagerVersion)

	   If iDiff > 0 Then
			sManagerVersion = sManagerVersion & String(iDiff, "0")
		Else
			sEngineVersion = sEngineVersion & String(iDiff, "0")	
	   End If

	   If CLng(sEngineVersion) > CLng(sManagerVersion)  Then
		   Reporter.ReportEvent MicFail, "Start Engine", "Failed Version Check" & vbCrLf & "Engine Expected " & Me.MinimunManagerVersion & _
																										vbCrLf & "But manager was actually " & Me.Manager.Version
		   CheckVersion = False
		   Exit Function
	   End If

	   'Check the engine version is good enough
	   sEngineVersion = Replace(CStr(Me.Version), ".", "")
	   sManagerVersion = Replace(CStr(Me.Manager.MinimumEngineVersion), ".","")

	   iDiff = Len(sEngineVersion) - Len(sManagerVersion)

	   If iDiff > 0 Then
			sManagerVersion = sManagerVersion & String(iDiff, "0")
		Else
			sEngineVersion = sEngineVersion & String(iDiff, "0")	
	   End If

	   If CLng(sEngineVersion) < CLng(sManagerVersion)  Then
		   Reporter.ReportEvent MicFail, "Start Engine", "Failed Version Check" & vbCrLf & "Manager expected " & Me.Manager.MinimumEngineVersion &_
																									  vbCrLf & "But engine actually was " & Me.Version
		   CheckVersion = False
		   Exit Function
	   End If

		CheckVersion = True
	End Function

	'*************************************************************'
	'	Sub Class_Initialize   (Private)
	'	Goal 								Setup Initialize event.
	'	In Params						None
	'	Out Params					None
	'*************************************************************'     
''' <summary>
''' Setup Initialize event
''' </summary>
   Private Sub Class_Initialize()   
	  Set oEngine = X 'Change X to the engine's object, 
										'or use any other method to create the engine
	  Init_Settings
	  'Add Initialization events here
   End Sub
   
	'*************************************************************'
	'	Sub Class_Terminate    (Private)
	'	Goal 								Setup Terminate event.
	'	In Params						None
	'	Out Params					None
	'*************************************************************'     
''' <summary>
''' Setup Terminate event
''' </summary>
	Private Sub Class_Terminate  
	   StopEngine
	   'Add Terminate Events Here
   End Sub

	'*************************************************************'
	'	Sub Load_Settings    (Private)
	'	Goal 								Loads sSettings into the oSettings object
	'	In Params						sSettings - The Filters init settings as "Field1>Value1|Field2>Value2|...
	'	Out Params					None
	'*************************************************************'  
	''' <summary>
	''' Loads sSettings into the oSettings object
	''' </summary>
	''' <param name="sSettings" type="String">The Engine's init settings as "Field1>Value1|Field2>Value2|...</param>
	Private Sub Load_Settings(sSettings) 
	   '***************************' No Need to change this sub! ***************************'
		''' <value type="Array"/>
		Dim ReplaceArray
		''' <value type="Array"/>
		Dim arrData, arrSettingData 
		''' <value type="Integer"/>
		Dim i

		'Protect >> and || chars
		ReplaceArray = Array ("ABCDEFGHIJKLMNOP1", "ABCDEFGHIJKLMNOP2")

			sSettings  = Replace(sSettings, "||", ReplaceArray(0))
			sSettings  = Replace(sSettings, ">>", ReplaceArray(1))
			
		   arrData = Split(sSettings, "|") 'Split to individual fields
	
			   For i = 0 To UBound(arrData) 'Loop through fields
				   If arrData(i) <> "" Then
						arrSettingData = Split(arrData(i), ">") 'Break to Field and to Value
						
						'Protect | char
						arrSettingData(0) = Replace(arrSettingData(0), ReplaceArray(0), "|")
						arrSettingData(1) = Replace(arrSettingData(1), ReplaceArray(0), "|")

						'Protect > char
						arrSettingData(0) = Replace(arrSettingData(0), ReplaceArray(1), ">")
						arrSettingData(1) = Replace(arrSettingData(1), ReplaceArray(1), ">")

						If oSettings.Exists(arrSettingData(0)) Then 'There's a default value for this field
							oSettings(arrSettingData(0)) = arrSettingData(1) 'Change it to the user's value
						Else
							oSettings.Add arrSettingData(0), arrSettingData(1) 'Create new setting
						End If
						
					End If
					
			   Next
	
	End Sub


		   
End Class
