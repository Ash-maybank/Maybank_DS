dd = day(date-1)
mm = month(date)
yy = year(date)
dt = yy*10000+mm*100+dd

'MsgBox("Welcome Day " & mid(dt,7,2))
'MsgBox("Welcome Mth " & mid(dt,5,2))
'MsgBox("Welcome Yr " & mid(dt,1,4))

'Const strFolder = "C:\Users\ps01072018\Desktop\RAW DATA\VBS\"
Const strFolder = "C:\Users\ps01072018\Desktop\RAW DATA\"
strFolderM = strFolder & mid(dt,1,6) & "\"
strFolderD = strFolderM & mid(dt,7,2) & "\"
Const strShared = "S:\Business Enablement & Compliance\005. Projects\002. Raw Data\"
strSharedM = strShared & mid(dt,1,6) & "\"
strSharedD = strSharedM & mid(dt,7,2) & "\"
Const Overwrite = True
'To create folder
Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")
'To copy folder
Dim objFileSys
Set objFileSys = CreateObject("Scripting.FileSystemObject")

' Create a new folder in PC for month
If Not oFSO.FolderExists(strFolderM) Then
  oFSO.CreateFolder strFolderM
End If
' Create a new folder in PC for day
If Not oFSO.FolderExists(strFolderD) Then
  oFSO.CreateFolder strFolderD
End If

' Create a new folder in Shared Folder for month
If Not oFSO.FolderExists(strSharedM) Then
  oFSO.CreateFolder strSharedM
End If
' Create a new folder in Shared Folder for day
If Not oFSO.FolderExists(strSharedD) Then
  oFSO.CreateFolder strSharedD
End If


' Copy a whole folder
objFileSys.GetFolder("X:\" & mid(dt,1,4) & "\" & mid(dt,5,2) & "\" & mid(dt,7,2) & "\00067\").Copy strFolder & mid(dt,1,6) & "\" & mid(dt,7,2) & "\"
objFileSys.GetFolder("X:\" & mid(dt,1,4) & "\" & mid(dt,5,2) & "\" & mid(dt,7,2) & "\00567\").Copy strFolder & mid(dt,1,6) & "\" & mid(dt,7,2) & "\"

objFileSys.GetFolder("X:\" & mid(dt,1,4) & "\" & mid(dt,5,2) & "\" & mid(dt,7,2) & "\00067\").Copy strSharedD
objFileSys.GetFolder("X:\" & mid(dt,1,4) & "\" & mid(dt,5,2) & "\" & mid(dt,7,2) & "\00567\").Copy strSharedD
