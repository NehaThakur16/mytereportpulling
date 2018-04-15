'Set objshell = WScript.CreateObject("WScript.Shell")
'key="HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\{374DE290-123F-4565-9164-39C4925E467B}"
'folder = objshell.RegRead(key)
'WScript.Echo folder
'objshell.Run folder


username=CreateObject("WScript.Network").UserName
MsgBox username

path1 = "C:\Users\" & username & "\Downloads\ Forecasted Time Details.xlsx"
pqath2 = ""
Set excelobj = CreateObject("Excel.Application")
Set wbobj1 = excelobj.Workbooks.Open(path1)
obj.Visible = True
Set wbobj2 = excelobj.Workbooks.Open(path2)


lastrow = wsobj2.Cells(Rows.count,1).End(xlUp).offset(1,0).Row
  
Set wsobj1 = Worksheet(1).Activate
wsobj1.Range("A1:f1").End(xlUp).Row.copy

Set wsobj2 = Worksheet(1).Activate
wsobj1.Paste Destination:= wsobj2.Cells(lastrow,1)
wbobj1.Close False
wbobj1.Delete

wsobj2.Save
wsobj2.Close
excelobj.Quit