Option Explicit
dim ObjFSO
set ObjFSO=WScript.CreateObject("Scripting.FileSystemObject")

'Driveオブジェクトとドライブ名を格納する変数宣言
Dim ObjDrive,StrDrive
StrDrive="\\10.1.100.203\d$"

if ObjFSO.DriveExists(StrDrive) then
	set ObjDrive=ObjFSO.GetDrive(StrDrive)
end if
if ObjDrive.IsReady then
	msgbox  FormatNumber(objDrive.FreeSpace / 1073741824, 2) & "GB"
else
	msgbox"ドライブが準備できていません"
end if  

set ObjFSO =nothing
set ObjDrive = nothing                                    
