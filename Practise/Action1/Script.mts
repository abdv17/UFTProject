

'Action on Webapplication
SystemUtil.Run "iexplore.exe","http://newtours.demoaut.com/mercurywelcome.php"
Dim sobjPage,sWebObject
Set sobjPage = Browser("Welcome: Mercury Tours").Page("Welcome: Mercury Tours")



Call fn_Web_UI_WebEdit_Opearations("WebEdit_Set","Set",sobjPage,"sg8412","userName")

Call fn_Web_UI_WebEdit_Opearations("WebEdit_Set","SetSecure",sobjPage,"594287779be76ea082631d7c53a20de41da4c8f38986f0f0","password")
Call fn_Web_UI_Image_Operations("Image_Click","Click",sobjPage,"LogIn")


Set sobjPage = Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury") @@ hightlight id_;_Browser("Welcome: Mercury Tours").Page("Welcome: Mercury Tours").WebEdit("password")_;_script infofile_;_ZIP::ssf6.xml_;_
Set sWebObject = sobjPage.WebTable("SIGN-OFF")
If fn_Web_UI_WebObject_Operations("WebObject_Exist","exist",sWebObject,"") Then
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