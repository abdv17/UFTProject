'Working with Excel Objects
'Writing functions to open an existing excel
'Read data,execute accordingly
'Update data after execution


sExcelFilePath = Environment.Value("TestDir") & "\TestData\TestCasesData.xlsx"

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWrkBk = objExcel.Workbooks.Open(sExcelFilePath)
Set objWrkSht = objWrkBk.Worksheets("Sheet1")

iRwcnt = objWrkSht.UsedRange.Rows.Count
iClncnt = objWrkSht.UsedRange.Columns.Count
For i = 1 To iRwcnt
	
		If objWrkSht.Cells(i,4).Value = "Yes" Then
			Environment.Value("bExecStatus") = False
			Environment.Value("ActionWord") = objWrkSht.Cells(i,3).Value
			Environment.Value("TCName") = objWrkSht.Cells(i,2).Value
			Call fn_CreateWordDoc()
			objWrkSht.Cells(i,6).Value = Time	
			LoadAndRunAction "C:\Users\Lenovo\Documents\Unified Functional Testing\Practise",Environment.Value("ActionWord"),oneIteration
			objWrkSht.Cells(i,7).Value = Time
			objWrkSht.Cells(i,8).Value = Second(objWrkSht.Cells(i,7).Value - objWrkSht.Cells(i,6).Value) & " Seconds"
			If Environment("bExecStatus") Then
				objWrkSht.Cells(i,5).Value = "Pass"
			Else
				objWrkSht.Cells(i,5).Value = "Fail"
			End If
		End If 
		SystemUtil.CloseProcessByName("iexplore.exe")
	
Next
objWrkBk.Save
objWrkBk.Close
Set objWrkSht = Nothing
Set objWrkBk = Nothing

objExcel.Quit




'Call for HTML reporting after test data excel sheet complete update done.
Call fn_HTMLReporting(sExcelFilePath)




'Code to send a mail with run results/report
Call fn_generateMail()
Public Function fn_generateMail()
	
'Code to send a mail with run results/report

Dim ToAddress 
Dim Subject 
Dim Body 
Dim Attachment 
Dim oUtlookApp, nAmeSpace, newMail 

SystemUtil.Run "C:\Program Files\Microsoft Office\Office11\OUTLOOK.EXE" 'This line should be enabled if the Outlook on the desktop is not running 

ToAddress = "sai.gaddam.ext@siemens.com" ' Message recipient Address 

Set oUtlookApp = CreateObject("Outlook.Application")

Set nAmeSpace = oUtlookApp.GetNamespace("MAPI") 

Set newMail = oUtlookApp.CreateItem(0) 

Subject = "This is a test mail" 'Message Subject you can update

Body = "This the message for testing. " 'Message body you can update

newMail.Subject = Subject 

newMail.Body = Body & vbCrLf 

newMail.Recipients.Add(ToAddress) 

newMail.Attachments.Add(Environment.Value("FunctionLogPath")) 'You can update attachment file name 
newMail.Attachments.Add(Environment.Value("TestDir") & "\Report.html")
newMail.Send 

Set nAmeSpace = Nothing 

Set oUtlookApp = Nothing 
End Function
