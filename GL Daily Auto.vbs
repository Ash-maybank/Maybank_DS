Dim ObjExcel, ObjWB
Set ObjExcel = CreateObject("excel.application")
'vbs opens a file specified by the path below
Set ObjWB = ObjExcel.Workbooks.Open("C:\Users\ps01072018\Desktop\Project\012. GL1200 and GL2460 Revenue\GL1200 and GL2460 Once v7 Daily Auto.xlsm")
'either use the Workbook Open event (if macros are enabled), or Application.Run

objExcel.Application.Visible = True

'ObjWB.Close False
ObjExcel.Quit
Set ObjExcel = Nothing