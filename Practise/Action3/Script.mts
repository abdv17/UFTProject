
'Environment.Value("bExecStatus") = True


LoadAndRunAction "C:\Users\Lenovo\Documents\Unified Functional Testing\Practise","Login",oneIteration
Environment.Value("bExecStatus") = False

Dim sobjPage,sWebObject
Set sobjPage = Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury")

Call fn_Web_UI_Link_Operations("Link_Click","Click",sobjPage,"ITINERARY")

Set sobjPage = Browser("Select a Flight: Mercury").Page("Select a Flight: Mercury")


Set sWebObject = sobjPage.Image("itinerary")



If fn_Web_UI_WebObject_Operations("ItineraryImage_Exist","exist",sWebObject,"") Then
	Call fn_PrintnUpdateLogFile("FunctionLogPath","<PASS>:[" & sWebObject.ToString & "] exists and Test Case Passed")
	Environment.Value("bExecStatus") = True
Else
	Call fn_PrintnUpdateLogFile("FunctionLogPath","<FAIL>:[" & sWebObject.ToString & "] doesn't exists and Test Case Failed")
End If


If Err.number <> 0 Then
	Set sobjPage = Nothing
	Set sWebObject = Nothing
	Call fn_PrintnUpdateLogFile("FunctionLogPath","<FAIL>:Test Case is Failed")	
	Call fn_ExitTest()
End If
Set sobjPage = Nothing
Set sWebObject = Nothing
Public Function fn_ExitTest()
	Set sobjPage = Nothing
	Set sWebObject = Nothing
	ExitTest
End Function
