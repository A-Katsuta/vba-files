@echo off
setlocal enabledelayedexpansion

:: -------------------------------------------------
:: ���@�\�t�@�C�����l�[���o�b�` (Ver.3)
:: �@�\: ���t�A�����A�J�X�^���������t�@�C�����ɒǉ�
:: -------------------------------------------------

:: --- �����ݒ� ---
set "TARGET_FOLDER=%cd%"  REM �Ώۃt�H���_ (�K��l: ���̃o�b�`�t�@�C��������ꏊ)
set "FILE_PATTERN=*.*"    REM �Ώۃt�@�C�� (�K��l: �S�Ẵt�@�C��)
set "RECURSIVE=/R"        REM �T�u�t�H���_��Ώۂɂ��邩 (/R) �󗓂Ȃ�ΏۊO

:MENU
cls
echo =================================================================
echo  ���@�\ �t�@�C�����ꊇ�ύX�c�[��
echo =================================================================
echo.
echo   ���݂̐ݒ�:
echo     �Ώۃt�H���_: %TARGET_FOLDER%
echo     �Ώۃt�@�C��: %FILE_PATTERN%
echo     �T�u�t�H���_: %RECURSIVE% ( /R=�܂� )
echo.
echo -----------------------------------------------------------------
echo.
echo  [ �ǉ�������̂�I�����Ă������� ]
echo.
echo    1. �t�@�C�����́y�擪�z�� [YYYYMMDD_] ��ǉ�
echo    2. �t�@�C�����́y�����z�� [_YYYYMMDD] ��ǉ�
echo    3. �t�@�C�����́y�����z�� [_YYYYMMDD-HHMMSS] ��ǉ�
echo.
echo    4. �t�@�C�����́y�擪�z�� [�J�X�^������_] ��ǉ�
echo    5. �t�@�C�����́y�����z�� [_�J�X�^������] ��ǉ� 
echo.
echo    8. [�e�X�g���s(Dry Run)] ���s�� (���O�̕ύX�͂��܂���)
echo    9. [�ݒ�̕ύX]
echo.
echo    Q. �I��
echo.
echo -----------------------------------------------------------------

set /p "CHOICE=�ԍ���I�����Ă�������: "

if /i "%CHOICE%"=="1" call :ProcessFiles "PREFIX_DATE" && goto MENU
if /i "%CHOICE%"=="2" call :ProcessFiles "SUFFIX_DATE" && goto MENU
if /i "%CHOICE%"=="3" call :ProcessFiles "SUFFIX_DATETIME" && goto MENU
if /i "%CHOICE%"=="4" call :ProcessFiles "PREFIX_CUSTOM" && goto MENU
if /i "%CHOICE%"=="5" call :ProcessFiles "SUFFIX_CUSTOM" && goto MENU
if /i "%CHOICE%"=="8" call :ProcessFiles "DRY_RUN" && goto MENU
if /i "%CHOICE%"=="9" goto SETTINGS
if /i "%CHOICE%"=="Q" goto EOF

echo �����ȑI���ł��B
pause
goto MENU


:SETTINGS
cls
echo.
echo --- �ݒ�̕ύX ---
echo   ���݂̐ݒ�l���\������Ă��܂��B�ύX���Ȃ��ꍇ�͂��̂܂�Enter�������Ă��������B
echo.
set /p "FILE_PATTERN=�Ώۃt�@�C���̃p�^�[������� (��: *.jpg, *.pptx, *.*) [%FILE_PATTERN%]: "
set /p "TARGET_FOLDER=�Ώۃt�H���_�̃p�X����͂��Ă������� [%TARGET_FOLDER%]: "
set /p "RECURSIVE_INPUT=�T�u�t�H���_���Ώۂɂ��܂����H (Y/N) : "
if /i "%RECURSIVE_INPUT%"=="Y" (set "RECURSIVE=/R") else (set "RECURSIVE=")
goto MENU


:ProcessFiles
cls
set "MODE=%~1"

rem --- ���t�Ǝ����̏����ݒ� (YYYYMMDD �� HHMMSS) ---
set "DATE_STR=%date:~0,4%%date:~5,2%%date:~8,2%"
set "TIME_STR=%time:~0,2%%time:~3,2%%time:~6,2%"
set "TIME_STR=%TIME_STR: =0%"

set "TIMESTAMP_DATE=%DATE_STR%"
set "TIMESTAMP_DATETIME=%DATE_STR%-%TIME_STR%"

rem --- �J�X�^�������̓��� ---
if "%MODE%"=="PREFIX_CUSTOM" or "%MODE%"=="SUFFIX_CUSTOM" (
    if "%MODE%"=="PREFIX_CUSTOM" set "POS=�擪"
    if "%MODE%"=="SUFFIX_CUSTOM" set "POS=����"
    set /p "CUSTOM_TEXT=�t�@�C������!POS!�ɒǉ����镶������͂��Ă�������: "
    if not defined CUSTOM_TEXT (
        echo ���������͂���܂���ł����B���j���[�ɖ߂�܂��B
        pause
        exit /b
    )
)

echo.
echo --- �������J�n���܂� ---
echo ���[�h: %MODE%
echo.

set "COUNT=0"

if "%MODE%"=="DRY_RUN" (
    echo ������ �e�X�g���s���[�h�ł� ������
    echo �ǂ̂悤�ɖ��O���ύX����邩��\�����܂��B���ۂ̕ύX�͍s���܂���B
    echo.
    echo [�ύX�O] ==^> [�ύX��]
    echo ----------------------------------
)

for %RECURSIVE% %%F in ("%TARGET_FOLDER%\%FILE_PATTERN%") do (
    set "FILE_NAME=%%~nF"
    set "FILE_EXT=%%~xF"
    set "NEW_NAME="

    rem --- ���[�h�ɉ����ĐV�����t�@�C�����𐶐� ---
    if "%MODE%"=="PREFIX_DATE" set "NEW_NAME=%TIMESTAMP_DATE%_!FILE_NAME!!FILE_EXT!"
    if "%MODE%"=="SUFFIX_DATE" set "NEW_NAME=!FILE_NAME!_%TIMESTAMP_DATE%!FILE_EXT!"
    if "%MODE%"=="SUFFIX_DATETIME" set "NEW_NAME=!FILE_NAME!_%TIMESTAMP_DATETIME%!FILE_EXT!"
    if "%MODE%"=="PREFIX_CUSTOM" set "NEW_NAME=%CUSTOM_TEXT%_!FILE_NAME!!FILE_EXT!"
    if "%MODE%"=="SUFFIX_CUSTOM" set "NEW_NAME=!FILE_NAME!_%CUSTOM_TEXT%!FILE_EXT!"

    if defined NEW_NAME (
        if "%MODE%"=="DRY_RUN" (
            echo "%%~nxF" ==^> "!NEW_NAME!"
        ) else (
            if not "%%~nxF"=="!NEW_NAME!" (
                ren "%%F" "!NEW_NAME!"
                echo RENAMED: "%%~nxF" --^> "!NEW_NAME!"
            )
        )
        set /a COUNT+=1
    )
)

echo.
echo ----------------------------------
echo �������������܂���: %COUNT% ���̃t�@�C�����������܂����B
echo.
pause
exit /b

:EOF
endlocal
exit