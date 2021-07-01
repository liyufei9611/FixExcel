:: 上海分行sftp腳本
:: 供數目標系統：上海分行
:: AUTHOR：TFB_李宇非 3334
:: DATE: 2021-06-28

::@ECHO OFF


SET NOWDAY=20210701
::SET NOWDAY=%date:~10,4%%date:~4,2%%date:~7,2%

set FILE_LIST_FILE=FILE.LIST

:: 讀取配置文件信息

For /f "tokens=1-2 delims==" %%i in (config\sftp_shb.ini) do (
	If "%%i"=="SFTP_IP" SET SFTP_IP=%%j
	If "%%i"=="SFTP_USER" SET SFTP_USER=%%j
	If "%%i"=="SFTP_PASSWORD" SET SFTP_PASSWORD=%%j
	If "%%i"=="SFTP_DATA_FOLDER" SET SFTP_DATA_FOLDER=%%j
	If "%%i"=="UNLOAD_DATA_FOLDER" SET UNLOAD_DATA_FOLDER=%%j
	If "%%i"=="LOG_HOME" SET LOG_HOME=%%j
	If "%%i"=="log_save_days" SET log_save_days=%%j		
)

set LOG_FILE=%LOG_HOME%\%NOWDAY%\sftp_shb_%NOWDAY%.log

:: 先创建當天日志文件夾
if not exist %LOG_HOME%\%NOWDAY% (
	mkdir %LOG_HOME%\%NOWDAY%)
	


:: 判斷是否有輸入日期參數，若有則賦值給變量BATCH_DATE

if not "%1"=="" (
	DEL /q %BATCH_DATE%
	set BATCH_DATE=%1
) else (
	set BATCH_DATE=%NOWDAY%
)

echo %BATCH_DATE%


echo ======Start sFTP data at %NOWDAY%======>>%LOG_FILE%
echo BATCH_DATE:%BATCH_DATE%======>>%LOG_FILE%


:: 卸數文件夾
SET CURR_DATA_FOLDER=%UNLOAD_DATA_FOLDER%\%BATCH_DATE%

echo %CURR_DATA_FOLDER%
echo %SFTP_IP%
echo %SFTP_USER%
echo %SFTP_PASSWORD%
echo %SFTP_DATA_FOLDER%


cd   %CURR_DATA_FOLDER%


::建立新的空文件
ECHO=>%unload_data_folder%\%BATCH_DATE%\FILE.LIST
::文件count
dir /b /a-d | find /v /c "" >>%FILE_LIST_FILE%

echo mkdir %SFTP_DATA_FOLDER%/%BATCH_DATE%>>sftp.txt
echo cd %SFTP_DATA_FOLDER%/%BATCH_DATE%>>sftp.txt
echo put %BATCH_DATE%.zip>>sftp.txt
echo bye>>sftp.txt

::压缩当前目录全部文件
"C:\Program Files\WinRAR\WinRAR.exe" a -r "%NOWDAY%.zip" -rr -m3 * -ibck

(
psftp %SFTP_IP% -l %SFTP_USER% -pw %SFTP_PASSWORD% -P 2022 -b sftp.txt
)>>%LOG_FILE%

del sftp.txt
del %NOWDAY%.zip



:: 刪除30天之前的日誌文件
::(forfiles /p %cd%\logs /s /m *.log /d -%log_save_days% /c "cmd /c del @path"
::) >>%LOG_FILE%

ECHO ======End sFTP data at %date% %time%======>>%LOG_FILE%

pause


:: cd $UNLOAD_DATA_FOLDER
:: rm -rf ${BATCH_DATE}
:: echo ======group 2 End SFTP data at $(date)======
