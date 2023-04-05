@echo off

net session >nul 2>&1
if %ERRORLEVEL% equ 0 (
    goto MAIN
   ) else (
    echo ■ 管理者モードではないため、処理を終了します ■
    echo    このbatを利用するためには、管理者として起動する必要があります。
    echo    このbatファイルを右クリック→「管理者として実行」をクリック
    pause
    goto END
    
 )
)

:MAIN  
cls

echo □■━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo ■　チューニングレベル（受信データ量レベル）の設定
echo ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo    ※このbatファイルを「管理者として実行」して、利用してください。
echo    コンピュータ本体のチューニングレベル（受信データ量レベル）を設定します。
echo    チューニングレベルが高ければ通信速度が向上するという事ではなく、
echo    コンピュータによって相性のよい設定値がありますので、
echo    それぞれの設定値をセットした後、速度テストを行い、相性の良い設定値に設定してください。
echo;
echo № 設定値              説明
echo 1  normal              初期設定
echo 2  highlyrestricted    既定値より受信ウインドウを若干拡大（推奨）
echo 3  restricted          既定値より受信ウインドウを拡大（推奨）
echo 4  experimental        既定値より受信ウインドウを大きく拡大
echo 5  disabled            自動最適化を無効
echo 6  処理終了
echo 7  設定状況を確認する

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
echo [normal](初期設定)を設定しました。
pause
goto MAIN

:HIG
netsh interface tcp set global autotuninglevel=highlyrestricted
echo [highlyrestricted](既定値より受信ウインドウを若干拡大（推奨）)を設定しました。
pause
goto MAIN

:RES
netsh interface tcp set global autotuninglevel=restricted
echo [restricted](既定値より受信ウインドウを拡大（推奨）)を設定しました。
pause
goto MAIN

:EXP
netsh interface tcp set global autotuninglevel=experimental
echo [experimental](既定値より受信ウインドウを大きく拡大)を設定しました。
pause
goto MAIN

:DIS
netsh interface tcp set global autotuninglevel=disabled
echo [disabled](自動最適化を無効)を設定しました。
pause
goto MAIN

:SHOW
echo 現在設定されているチューニングレベルの状況は以下の通りです
netsh interface tcp show global
pause
goto MAIN


:END
exit/b