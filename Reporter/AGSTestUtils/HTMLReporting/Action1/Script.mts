 'Launching the application
strUrl="http://newtours.demoaut.com"
SetProject "Mercury Tours","Reservation","C:\temp\"

LogMessage gStart ,"Opening Application ","URL: " & strURL

LogMessage gInfo,"Launching Applcation", "URL: "& strURL
SystemUtil.CloseProcessByName "iexplore.exe"
Systemutil.Run "iexplore",strURL

LogMessage gStart ,"Starting  Login","Logging into Application"
LogMessage gInfo ,"Enter User Name :","mercury"
Browser(objBrowser).Page(objWelcomePage).WebEdit(user).Set "mercury"


LogMessage gInfo ,"Enter Password:","mercury"
Browser(objBrowser).Page(objWelcomePage).WebEdit(pwd).Set "mercury"

LogMessage gInfo ,"Clicking Login ",""
Browser(objBrowser).Page(objWelcomePage).Image(img).Click
Browser(objBrowser).Sync
Browser(objBrowser).Page(objWelcomePage).Sync

If Browser(objBrowser).Page(objSearchPage).Image(img1).Exist then
               LogMessage gPass,"Validating Login", "User  Logged in Successfully"
Else
               LogMessage gfail,"Validating Login", "Error in Login.Unable to Prcoeed Further"
                ExitTest
End If
LogMessage gStart ,"Starting  of Create A Flight Booking","Selecting Flight Details"
LogMessage gInfo ,"Selecting Trip Type :","Round Trip"
Browser(objBrowser).Page(objSearchPage).WebRadioGroup(objrdb).Click 

LogMessage gInfo ,"Selecting Departing From: :","Fankfurt"
Browser(objBrowser).Page(objSearchPage).WebList(pas).Select "2"

LogMessage gInfo ,"Selecting Arriving In: :","London"
Browser(objBrowser).Page(objSearchPage).WebList(objFP).Select "London"

LogMessage gInfo ,"Selecting Flying on :","July 15"
Browser(objBrowser).Page(objSearchPage).WebList(objFM).Select "July"
Browser(objBrowser).Page(objSearchPage).WebList(objFD).Select "15"

LogMessage gInfo ,"Selecting Arriving In :","Paris"
Browser(objBrowser).Page(objSearchPage).WebList(objTP).Select "Paris"

LogMessage gInfo ,"Selecting Returning on :","August,25"
Browser(objBrowser).Page(objSearchPage).WebList(objTM).Select "August"
Browser(objBrowser).Page(objSearchPage).WebList(objTD).Select "25"

LogMessage gInfo ,"Selecting Service Class :","Economy class "
Browser(objBrowser).Page(objSearchPage).WebRadioGroup(objSer).Click

LogMessage gInfo ,"Selecting Airline :","Unified Airlines"
Browser(objBrowser).Page(objSearchPage).WebList(Air).Select "Unified Airlines"

LogMessage gInfo ,"Clicking Continue :",""
Browser(objBrowser).Page(objSearchPage).Image(img1).Click
Browser(objBrowser).Sync
Browser(objBrowser).Page(objSearchPage).Sync
 
If Browser(objBrowser).page(objSelectPage).Image(img2).Exist  Then
			LogMessage gPass,"Validating Search Flights", "Flights search  Successful"
Else
               LogMessage gfail,"Validating SearchFlights", "Error in FlightsFind.Unable to Proceed Further"
                ExitTest
End If
LogMessage gStart ,"Starting  of Selecting Flight",""
Browser(objBrowser).page(objSelectPage).WebRadioGroup(rdbDep).Click

LogMessage gInfo ,"Selecting Return Flight",""
Browser(objBrowser).page(objSelectPage).WebRadioGroup(rdbRet).Click
LogMessage gInfo ,"Clicking Continue :",""
Browser(objBrowser).page(objSelectPage).Image(img2).Click
Browser(objBrowser).Sync
Browser(objBrowser).page(objSelectPage).Sync
If   Browser(objBrowser).page(objBookPage).WebEdit(objFN).Exist Then
   LogMessage gPass,"Validating FlightBooking", "Flight reservation is successful"
Else
   LogMessage gfail,"Validating FlightBooking", "Error in Booking Flights.Unable to Proceed Further"
   ExitTest
End If
LogMessage gStart ,"Starting  of Passenger Details","Entering Passenger Details"
LogMessage gStart ,"Entering Firstname","Adam"
Browser(objBrowser).page(objBookPage).WebEdit(objFN).Set "Adam"

LogMessage gInfo ,"Entering Firstname","MR"
Browser(objBrowser).page(objBookPage).WebEdit(objLN).Set "Mr"

LogMessage gInfo ,"Entering Card No","Adam"

Browser(objBrowser).page(objBookPage).WebEdit(objCN).Set "4567123412341234"
Browser(objBrowser).page(objBookPage).WebCheckBox(objchk1).Set "ON"

LogMessage gInfo ,"Clicking  Secure  Purchase",""
Browser(objBrowser).page(objBookPage).Image(img3).Click
Browser(objBrowser).Sync
Browser(objBrowser).page(objBookPage).Sync
If Browser(objBrowser).page(objConformPage).Image("file name:=backtoflights.gif").Exist  Then
	   LogMessage gPass,"Validating Confirmation", "Confirmation Successfull"
Else
      LogMessage gfail,"Validating Confirmation", "Error in Confirmation.Unable to Prcoeed Further"
      ExitTest
End If
LogMessage gInfo ,"Clicking Logout",""
Browser(objBrowser).Page(objConformPage).Link(objLogout).Click

Summarize_Results "MercuryTours_Reservation_001","MercuryTours_Reservations1","MercuryTours_Reservations1"
