option explicit
Dim fs
Dim fn
Dim renfile as string
Dim fname_before as string
dim fname_after as string

set fs = wscript.createobject("scripting.filesystemobject")
fname_before = "C:\Users\k-kasuga\AppData\Roaming\Microsoft\Excel\XLSTART\�}�N���܂Ƃ߃u�b�N(20190820).xlsm"
'�}�N���u�b�N���݃`�F�b�N
if fso.fileexists("C:\Users\k-kasuga\AppData\Roaming\Microsoft\Excel\XLSTART\�}�N���܂Ƃ߃u�b�N(20190820).xlsm")=false then
   msgbox"�ړ��Ώۂ̃t�@�C�������݂��܂���",vbexclamation
   goto ext
end if

'���l�[����t�@�C���d���`�F�b�N
�f���l�[����t�@�C�����擾

'���t�H���_



set fn = fs.getfile("C:\Users\k-kasuga\AppData\Roaming\Microsoft\Excel\XLSTART\�}�N���܂Ƃ߃u�b�N(20190820).xlsm")

fn.name = "�}�N���܂Ƃ߃u�b�N(20190820)_20210413.xlsm"

set fs = nothing
set fn = nothing