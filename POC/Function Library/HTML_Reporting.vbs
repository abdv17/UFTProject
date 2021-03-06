
'Function to report the execution status in HTML format
'Data is gather from TestData excel sheet after complete update of excel file
Public Function fn_HTMLReporting(sExcelFilePath)
	Set ofso = CreateObject("Scripting.FileSystemObject")
	Set ftxt = ofso.CreateTextFile (Environment.Value("TestDir") & "\Report.html")
	ftxt.WriteLine "<html>"
	ftxt.WriteLine "<head>"
	ftxt.WriteLine "<title>Report</title>"
	ftxt.WriteLine "<style>"
	ftxt.WriteLine ".Fail{"
	ftxt.WriteLine "background-color : Red;"
	ftxt.WriteLine "text-align : Center"
	ftxt.WriteLine "}"
	ftxt.WriteLine ".Pass{"
	ftxt.WriteLine "background-color : Green;"
	ftxt.WriteLine "text-align : Center"
	ftxt.WriteLine "}"
	ftxt.WriteLine "</style>"
	ftxt.WriteLine "</head>"
	ftxt.WriteLine "<body>"
	ftxt.WriteLine "<h1>This is a report of test run.</h1>"
	ftxt.WriteLine "<table border = 1px style=width:100%>"
	ftxt.WriteLine "<tbody>"
	ftxt.WriteLine "<tr>"
	ftxt.WriteLine "<th>SNo</th>"
	ftxt.WriteLine "<th>Name</th>"
	ftxt.WriteLine "<th>Start Time</th>"
	ftxt.WriteLine "<th>End Time</th>"
	ftxt.WriteLine "<th>Status</th>"
	ftxt.WriteLine "<th>Duration</th>"
	ftxt.WriteLine "</tr>"
	
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False
	Set objWrkBk = objExcel.Workbooks.Open(sExcelFilePath)
	Set objWrkSht = objWrkBk.Worksheets("Sheet1")
	iRwcnt = objWrkSht.usedrange.rows.count
	iClncnt = objWrkSht.usedrange.columns.count
	
	For i = 1 To iRwcnt
		If objWrkSht.Cells(i,4).Value = "Yes" Then
			'dynamic row of html starts here
			ftxt.WriteLine "<tr>"
			ftxt.WriteLine "<td>" & objWrkSht.cells(i,1).Value & "</td>"
			ftxt.WriteLine "<td>" & objWrkSht.Cells(i,2).Value & "</td>"
			ftxt.WriteLine "<td>" & CDate(objWrkSht.Cells(i,6).Value) &"</td>"
			ftxt.WriteLine "<td>" & CDate(objWrkSht.Cells(i,7).Value) & "</td>"
			ftxt.WriteLine "<td class=" & objWrkSht.Cells(i,5).Value & ">" & objWrkSht.Cells(i,5).Value & "</td>"
			ftxt.WriteLine "<td>" & objWrkSht.Cells(i,8).Value & "</td>"
			ftxt.WriteLine "</tr>"
		End If
	Next
	
	objWrkBk.Close
	Set objWrkSht = Nothing
	Set objWrkBk = Nothing
	
	objExcel.Quit
	ftxt.WriteLine "</tbody>"
	ftxt.WriteLine "</table>"
	ftxt.WriteLine "</body>"
	ftxt.WriteLine "</html>"
	ftxt.Close
	Set ftxt = Nothing
	Set ofso = Nothing	
End Function