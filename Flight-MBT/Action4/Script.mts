﻿' Just close the window since there isn't a logout option
' This action can occur when any main screen is displayed (without any dependant popup windows).

If WpfWindow("OpenText MyFlight Sample Application").Exist Then
	WpfWindow("OpenText MyFlight Sample Application").Close
Else
	Reporter.ReportEvent micFail, "Logout - Flight GUI not found", "Flight GUI Window not found. Nothing to close. Exiting..."
End If
 @@ hightlight id_;_1705708_;_script infofile_;_ZIP::ssf7.xml_;_
