Option Explicit
call FileCountCheck

Sub FileCountCheck()
Dim fso, flg, X, fl,folderpath,filepath
dim i
dim folderO(10)
dim folderP(10)
dim folderCMAX(10)
dim folderC(10)


if msgbox("��ʁE���[�����i�[�̊e�t�H���_���̃t�@�C�������`�F�b�N���܂��B" & vbcrlf & "���s���Ă�낵���ł����H",vbyesno + vbinformation) <> vbyes then
exit sub
end if 

'�h���C�����[�U�[�����擾���Ak-kasuga��k-satou�ȊO�A������ʂ��Ȃ�
Set X = WScript.CreateObject("WScript.network")
If x.username <> "k-kasuga" and x.username <> "k-satou" Then
   MsgBox "k-kasuga�[���܂��́Ak-satou�[���̂ݎ��s�\�ł�",vbcritical
   Set X = Nothing
  Exit Sub
End If


Set fso = CreateObject("Scripting.FileSystemObject")

'�J�����g�t�H���_�̃p�X���擾���A�w�肷��ꍇ�͏��folderpath�擾���J��
'��΃p�X�Ŏw�肷��ꍇ�͉���folderpath�擾���J������
'folderpath = fso.getParentFolderName(WScript.ScriptFullName)�@
'folderpath = "C:\Users\k-kasuga\Desktop\910_��ʁE���[�����i�["           
folderpath = "\\KPDSV1\Share\���蕔��G���A\00_�����V�X�e��\910_��ʁE���[�����i�["           


'�`�F�b�N����e�t�H���_�̃p�X���`
folderP(1) =folderpath & "\00_���̑��E����"
folderP(2) =folderpath & "\01_�̔�"
folderP(3) =folderpath & "\02_���B"
folderP(4) =folderpath & "\03_�݌�"
folderP(5) =folderpath & "\04_���Y"
folderP(6) =folderpath & "\05_����"
folderP(7) =folderpath & "\06_��"
folderP(8) =folderpath & "\07_�\�Z"
folderP(9) =folderpath & "\08_�����E����"
folderP(10) =folderpath & "\09_�}�X�^"         


'�`�F�b�N����e�t�H���_�̍ő�t�@�C�������`
folderCMAX(1) = 1
folderCMAX(2) = 30
folderCMAX(3) = 21
folderCMAX(4) = 34
folderCMAX(5) = 12
folderCMAX(6) = 5
folderCMAX(7) = 24
folderCMAX(8) = 14
folderCMAX(9) = 33
folderCMAX(10) = 49


'�e�t�H���_�̃I�u�W�F�N�g���擾
for i = 1 to 10
	set folderO(i) = fso.getfolder(folderP(i))
next 


'�e�t�H���_���̃t�@�C�������擾
for i = 1 to 10
	folderC(i) = folderO(i).files.count
next   

'msgbox�Ƀt�@�C�����̔�r���ʂ�\��
msgbox "�y00_���̑��E���ʁz" & folderC(1) & " / " & folderCMAX(1)& vbcrlf & _
       "�y01_�̔��z" & folderC(2) & " / " & folderCMAX(2)& vbcrlf & _
       "�y02_���B�z" & folderC(3) & " / " & folderCMAX(3)& vbcrlf & _
       "�y03_�݌Ɂz" & folderC(4) & " / " & folderCMAX(4)& vbcrlf & _
       "�y04_���Y�z" & folderC(5) & " / " & folderCMAX(5)& vbcrlf & _
       "�y05_�����z" & folderC(6) & " / " & folderCMAX(6)& vbcrlf & _
       "�y06_���z" & folderC(7) & " / " & folderCMAX(7)& vbcrlf & _
       "�y07_�\�Z�z" & folderC(8) & " / " & folderCMAX(8)& vbcrlf & _
       "�y08_�����E�����z" & folderC(9) & " / " & folderCMAX(9)& vbcrlf & _
       "�y09_�}�X�^�z" & folderC(10) & " / " & folderCMAX(10)
        
'�I�u�W�F�N�g�̔p��
Set fso = Nothing

End Sub