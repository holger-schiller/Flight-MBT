﻿' Search for an existing order by one of three values which are passed in (only one should be non-empty)
'This action must start on the post-login screen that shows BOOK FLIGHT and SEARCH ORDER.

a=Parameter("OrderNumber")
b=Parameter("OrderDate")
c=Parameter("PassengerName")

'Get to BOOK FLIGHT and SEARCH ORDER screen if you aren't already there
If WpfWindow("OpenText MyFlight Sample Application").WpfButton("NEW SEARCH").Exist (4) Then @@ hightlight id_;_2137773064_;_script infofile_;_ZIP::ssf9.xml_;_
	WpfWindow("OpenText MyFlight Sample Application").WpfButton("NEW SEARCH").Click
End If

' If we're not on the SEARCH ORDER screen then exit
If WpfWindow("OpenText MyFlight Sample Application").WpfTabStrip("WpfTabStrip").Exist Then
	WpfWindow("OpenText MyFlight Sample Application").WpfTabStrip("WpfTabStrip").Select "SEARCH ORDER"
Else
	Reporter.ReportEvent micFail, "Search Order - Wrong screen", "Not on SEARCH ORDER screen. Exiting..."
	ExitAction
End If

' Only pass in one of these three values - only one should be non-empty
If a <> "" Then ' Process by Order Number
	WpfWindow("OpenText MyFlight Sample Application").WpfRadioButton("byNameOrDateRadio").Set
	WpfWindow("OpenText MyFlight Sample Application").WpfEdit("byNameWatermark").Set " " ' Clear out field
	WpfWindow("OpenText MyFlight Sample Application").WpfRadioButton("byNumberRadio").Set @@ hightlight id_;_2062105976_;_script infofile_;_ZIP::ssf3.xml_;_
	WpfWindow("OpenText MyFlight Sample Application").WpfEdit("byNumberWatermark").Set a
	WpfWindow("OpenText MyFlight Sample Application").WpfButton("SEARCH").Click
ElseIf b <>  "" Then ' Process by Order Date
	WpfWindow("OpenText MyFlight Sample Application").WpfRadioButton("byNameOrDateRadio").Set
	WpfWindow("OpenText MyFlight Sample Application").WpfEdit("byNameWatermark").Set "" ' Clear out field
	WpfWindow("OpenText MyFlight Sample Application").WpfCalendar("byDatePicker").SetDate b @@ hightlight id_;_1955569168_;_script infofile_;_ZIP::ssf15.xml_;_
	WpfWindow("OpenText MyFlight Sample Application").WpfButton("SEARCH").Click
	rc = WpfWindow("OpenText MyFlight Sample Application").WpfTable("ordersDataGrid").RowCount
	If rc > 0 Then
		WpfWindow("OpenText MyFlight Sample Application").WpfTable("ordersDataGrid").SelectRow rc-1 ' Select the last row (zero-based)
		WpfWindow("OpenText MyFlight Sample Application").WpfButton("SELECT ORDER").Click
	Else
		Reporter.ReportEvent micFail, "Search Order - Date not found", "Order with Order Date " & b & " does not exist. Exiting..."
		WpfWindow("OpenText MyFlight Sample Application").WpfButton("BACK").Click
		ExitAction
	End If
ElseIf c <> "" Then ' Process by Passenger Name
	WpfWindow("OpenText MyFlight Sample Application").WpfRadioButton("byNameOrDateRadio").Set @@ hightlight id_;_1918806040_;_script infofile_;_ZIP::ssf30.xml_;_
	WpfWindow("OpenText MyFlight Sample Application").WpfCalendar("byDatePicker").Type micBack
	WpfWindow("OpenText MyFlight Sample Application").WpfEdit("byNameWatermark").Set c
	WpfWindow("OpenText MyFlight Sample Application").WpfButton("SEARCH").Click
	rc = WpfWindow("OpenText MyFlight Sample Application").WpfTable("ordersDataGrid").RowCount
	If rc > 0 Then
		WpfWindow("OpenText MyFlight Sample Application").WpfTable("ordersDataGrid").SelectRow rc-1 ' Select the last row (zero-based)
		WpfWindow("OpenText MyFlight Sample Application").WpfButton("SELECT ORDER").Click
	Else
		Reporter.ReportEvent micFail, "Search Order - Name not found", "Order with Passenger Name containing " & c & " does not exist. Exiting..."
		WpfWindow("OpenText MyFlight Sample Application").WpfButton("BACK").Click
		ExitAction
	End If	
End If

If WpfWindow("OpenText MyFlight Sample Application").Dialog("Error").Exist (2) Then ' This popup occurs only with non-existant order nums @@ hightlight id_;_2137773064_;_script infofile_;_ZIP::ssf9.xml_;_
	WpfWindow("OpenText MyFlight Sample Application").Dialog("Error").WinButton("OK").Click
	Reporter.ReportEvent micFail, "Search Order - Number not found", "Order with Order Number " & a & " does not exist. Exiting..."
End If
 @@ hightlight id_;_853570_;_script infofile_;_ZIP::ssf24.xml_;_
' Flight GUI app ends on ORDER DETAILS screen with NEW SEARCH  button and the Trashcan icon available
' This action can then transition to either Change Order or Delete Order (or Logout).

