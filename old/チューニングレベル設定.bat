@echo off

net session >nul 2>&1
if %ERRORLEVEL% equ 0 (
    goto MAIN
   ) else (
    echo �� �Ǘ��҃��[�h�ł͂Ȃ����߁A�������I�����܂� ��
    echo    ����bat�𗘗p���邽�߂ɂ́A�Ǘ��҂Ƃ��ċN������K�v������܂��B
    echo    ����bat�t�@�C�����E�N���b�N���u�Ǘ��҂Ƃ��Ď��s�v���N���b�N
    pause
    goto END
    
 )
)

:MAIN  
cls

echo ����������������������������������������������������������������������������������������������������������������������������������������������������������
echo ���@�`���[�j���O���x���i��M�f�[�^�ʃ��x���j�̐ݒ�
echo ����������������������������������������������������������������������������������������������������������������������������������������������������������
echo    ������bat�t�@�C�����u�Ǘ��҂Ƃ��Ď��s�v���āA���p���Ă��������B
echo    �R���s���[�^�{�̂̃`���[�j���O���x���i��M�f�[�^�ʃ��x���j��ݒ肵�܂��B
echo    �`���[�j���O���x����������ΒʐM���x�����シ��Ƃ������ł͂Ȃ��A
echo    �R���s���[�^�ɂ���đ����̂悢�ݒ�l������܂��̂ŁA
echo    ���ꂼ��̐ݒ�l���Z�b�g������A���x�e�X�g���s���A�����̗ǂ��ݒ�l�ɐݒ肵�Ă��������B
echo;
echo �� �ݒ�l              ����
echo 1  normal              �����ݒ�
echo 2  highlyrestricted    ����l����M�E�C���h�E���኱�g��i�����j
echo 3  restricted          ����l����M�E�C���h�E���g��i�����j
echo 4  experimental        ����l����M�E�C���h�E��傫���g��
echo 5  disabled            �����œK���𖳌�
echo 6  �����I��
echo 7  �ݒ�󋵂��m�F����

choice /c 1234567       
if errorlevel 7 goto SHOW
if errorlevel 6 goto END
if errorlevel 5 goto DIS
if errorlevel 4 goto EXP
if errorlevel 3 goto RES
if errorlevel 2 goto HIG
if errorlevel 1 goto NOR
     
:NOR
netsh interface tcp set global autotuninglevel=normal
echo [normal](�����ݒ�)��ݒ肵�܂����B
pause
goto MAIN

:HIG
netsh interface tcp set global autotuninglevel=highlyrestricted
echo [highlyrestricted](����l����M�E�C���h�E���኱�g��i�����j)��ݒ肵�܂����B
pause
goto MAIN

:RES
netsh interface tcp set global autotuninglevel=restricted
echo [restricted](����l����M�E�C���h�E���g��i�����j)��ݒ肵�܂����B
pause
goto MAIN

:EXP
netsh interface tcp set global autotuninglevel=experimental
echo [experimental](����l����M�E�C���h�E��傫���g��)��ݒ肵�܂����B
pause
goto MAIN

:DIS
netsh interface tcp set global autotuninglevel=disabled
echo [disabled](�����œK���𖳌�)��ݒ肵�܂����B
pause
goto MAIN

:SHOW
echo ���ݐݒ肳��Ă���`���[�j���O���x���̏󋵂͈ȉ��̒ʂ�ł�
netsh interface tcp show global
pause
goto MAIN


:END
exit/b