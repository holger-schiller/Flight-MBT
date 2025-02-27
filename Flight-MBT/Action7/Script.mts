﻿' Change an existing flight according to the passed in 3 parameters
' Once at least one item is changed, the Check icon (update button) is enabled as shown in orange
' This action must start on the screen that shows the ORDER DETAILS.
'Typically that means you must run a SEARCH ORDER action first.
 
a=Parameter("Tickets")
b=Parameter("ClassOfService")
c=Parameter("PassengerName")

' If we're not on the ORDER DETAILS screen then exit
If NOT WpfWindow("OpenText MyFlight Sample Application").WpfObject("ORDER DETAILS").Exist Then
	Reporter.ReportEvent micFail, "Change Order - Wrong Screen", "Not on ORDER DETAILS screen. Exiting..."
	ExitAction
End  If

WpfWindow("OpenText MyFlight Sample Application").WpfComboBox("numOfTicketsCombo").Select a @@ hightlight id_;_1904552888_;_script infofile_;_ZIP::ssf4.xml_;_
WpfWindow("OpenText MyFlight Sample Application").WpfComboBox("flightClassCombo").Select b @@ hightlight id_;_1955290176_;_script infofile_;_ZIP::ssf8.xml_;_
WpfWindow("OpenText MyFlight Sample Application").WpfEdit("passengerName").Set c @@ hightlight id_;_1955294928_;_script infofile_;_ZIP::ssf10.xml_;_

'Make sure we change something or we can't click on the update button since it is disabled.
If WpfWindow("OpenText MyFlight Sample Application").WpfButton("updateBtn").GetROProperty("enabled")  = True Then
	WpfWindow("OpenText MyFlight Sample Application").WpfButton("updateBtn").Click
Else
	WpfWindow("OpenText MyFlight Sample Application").WpfComboBox("numOfTicketsCombo").Select 4 ' we set to 2 so 4 should be a different value
	WpfWindow("OpenText MyFlight Sample Application").WpfButton("updateBtn").Click ' the Check icon should now be enabled
End  If

Wait 2 ' The Exist statement below seems to find a little box that exists if we run too fast - so the Wait is necessary
' Wait for Update message to appear (about 3-4 seconds)
If WpfWindow("OpenText MyFlight Sample Application").WpfObject("OrderUpdatedMessage").Exist (4) Then
	WpfWindow("OpenText MyFlight Sample Application").WpfObject("OrderUpdatedMessage").Output CheckPoint("OrderUpdatedNumber") @@ hightlight id_;_1939452064_;_script infofile_;_ZIP::ssf19.xml_;_
End  If

' Flight GUI app ends on the Order Details confirmation screen showing the NEW SEARCH button.
