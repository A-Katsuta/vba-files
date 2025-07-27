@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

:: ==========================================================
:: ���@�\�t�@�C�����ꊇ�u���o�b�`
:: �@�\: ���K�\���Ή��A�v���r���[�A���O�o�́A���S���`�F�b�N
:: ==========================================================

set "VERSION=1.0"
set "SCRIPT_NAME=FileRenamer"

:: �J���[�ݒ�
set "C_RESET=[0m"
set "C_RED=[91m"
set "C_GREEN=[92m"
set "C_YELLOW=[93m"
set "C_BLUE=[94m"
set "C_MAGENTA=[95m"
set "C_CYAN=[96m"
set "C_WHITE=[97m"

:: �����ݒ�
set "TARGET_DIR=%CD%"
set "SEARCH_PATTERN="
set "REPLACE_PATTERN="
set "PREVIEW_MODE=1"
set "USE_REGEX=0"
set "CASE_SENSITIVE=0"
set "INCLUDE_SUBDIRS=0"
set "FILE_FILTER=*.*"
set "LOG_FILE="
set "BACKUP_MODE=0"

:: ���C�������J�n
:MAIN
cls
call :SHOW_HEADER
echo.

:: �p�����[�^������ꍇ�͒��ڎ��s���[�h
if not "%~1"=="" (
    call :PARSE_ARGS %*
    goto :EXECUTE_RENAME
)

:: �C���^���N�e�B�u���[�h
call :INTERACTIVE_MODE
goto :END

:: ========== �T�u���[�`�� ==========

:SHOW_HEADER
echo %C_CYAN%===============================================%C_RESET%
echo %C_CYAN%  %SCRIPT_NAME% v%VERSION%  %C_RESET%
echo %C_CYAN%  ���@�\�t�@�C�����ꊇ�u���c�[��%C_RESET%
echo %C_CYAN%===============================================%C_RESET%
goto :EOF

:INTERACTIVE_MODE
echo %C_WHITE%���݂̐ݒ�:%C_RESET%
echo   �Ώۃf�B���N�g��: %C_YELLOW%!TARGET_DIR!%C_RESET%
echo   �����p�^�[��: %C_YELLOW%!SEARCH_PATTERN!%C_RESET%
echo   �u���p�^�[��: %C_YELLOW%!REPLACE_PATTERN!%C_RESET%
echo   �t�@�C���t�B���^: %C_YELLOW%!FILE_FILTER!%C_RESET%
echo   �v���r���[���[�h: %C_YELLOW%!PREVIEW_MODE!%C_RESET%
echo   ���K�\���g�p: %C_YELLOW%!USE_REGEX!%C_RESET%
echo   �啶�����������: %C_YELLOW%!CASE_SENSITIVE!%C_RESET%
echo   �T�u�f�B���N�g���܂�: %C_YELLOW%!INCLUDE_SUBDIRS!%C_RESET%
echo   �o�b�N�A�b�v���[�h: %C_YELLOW%!BACKUP_MODE!%C_RESET%
echo.

echo %C_GREEN%���j���[:%C_RESET%
echo   %C_WHITE%1%C_RESET% - �Ώۃf�B���N�g����ݒ�
echo   %C_WHITE%2%C_RESET% - �����p�^�[����ݒ�
echo   %C_WHITE%3%C_RESET% - �u���p�^�[����ݒ�
echo   %C_WHITE%4%C_RESET% - �t�@�C���t�B���^��ݒ�
echo   %C_WHITE%5%C_RESET% - �I�v�V�����ݒ�
echo   %C_WHITE%6%C_RESET% - �v���r���[���s
echo   %C_WHITE%7%C_RESET% - ���s�i���ۂɃ��l�[���j
echo   %C_WHITE%8%C_RESET% - �w���v�\��
echo   %C_WHITE%0%C_RESET% - �I��
echo.

set /p "CHOICE=�I�����Ă������� (0-8): "

if "%CHOICE%"=="1" call :SET_TARGET_DIR
if "%CHOICE%"=="2" call :SET_SEARCH_PATTERN
if "%CHOICE%"=="3" call :SET_REPLACE_PATTERN
if "%CHOICE%"=="4" call :SET_FILE_FILTER
if "%CHOICE%"=="5" call :SET_OPTIONS
if "%CHOICE%"=="6" call :PREVIEW_RENAME
if "%CHOICE%"=="7" call :EXECUTE_RENAME
if "%CHOICE%"=="8" call :SHOW_HELP
if "%CHOICE%"=="0" goto :END

goto :INTERACTIVE_MODE

:SET_TARGET_DIR
echo.
echo %C_CYAN%�Ώۃf�B���N�g���̐ݒ�%C_RESET%
set /p "NEW_DIR=�f�B���N�g���p�X����� (����: !TARGET_DIR!): "
if not "!NEW_DIR!"=="" (
    if exist "!NEW_DIR!" (
        set "TARGET_DIR=!NEW_DIR!"
        echo %C_GREEN%�ݒ芮��: !TARGET_DIR!%C_RESET%
    ) else (
        echo %C_RED%�G���[: �w�肳�ꂽ�f�B���N�g�������݂��܂���%C_RESET%
    )
)
pause
goto :EOF

:SET_SEARCH_PATTERN
echo.
echo %C_CYAN%�����p�^�[���̐ݒ�%C_RESET%
set /p "NEW_PATTERN=�����p�^�[������� (����: !SEARCH_PATTERN!): "
if not "!NEW_PATTERN!"=="" set "SEARCH_PATTERN=!NEW_PATTERN!"
echo %C_GREEN%�ݒ芮��: !SEARCH_PATTERN!%C_RESET%
pause
goto :EOF

:SET_REPLACE_PATTERN
echo.
echo %C_CYAN%�u���p�^�[���̐ݒ�%C_RESET%
set /p "NEW_PATTERN=�u���p�^�[������� (����: !REPLACE_PATTERN!): "
set "REPLACE_PATTERN=!NEW_PATTERN!"
echo %C_GREEN%�ݒ芮��: !REPLACE_PATTERN!%C_RESET%
pause
goto :EOF

:SET_FILE_FILTER
echo.
echo %C_CYAN%�t�@�C���t�B���^�̐ݒ�%C_RESET%
echo ��: *.txt, *.jpg, *.*, IMG_*.jpg
set /p "NEW_FILTER=�t�B���^����� (����: !FILE_FILTER!): "
if not "!NEW_FILTER!"=="" set "FILE_FILTER=!NEW_FILTER!"
echo %C_GREEN%�ݒ芮��: !FILE_FILTER!%C_RESET%
pause
goto :EOF

:SET_OPTIONS
echo.
echo %C_CYAN%�I�v�V�����ݒ�%C_RESET%
echo   %C_WHITE%1%C_RESET% - ���K�\���g�p (����: !USE_REGEX!)
echo   %C_WHITE%2%C_RESET% - �啶����������� (����: !CASE_SENSITIVE!)
echo   %C_WHITE%3%C_RESET% - �T�u�f�B���N�g���܂� (����: !INCLUDE_SUBDIRS!)
echo   %C_WHITE%4%C_RESET% - �o�b�N�A�b�v���[�h (����: !BACKUP_MODE!)
echo   %C_WHITE%5%C_RESET% - ���O�t�@�C���ݒ� (����: !LOG_FILE!)
echo   %C_WHITE%0%C_RESET% - �߂�
echo.

set /p "OPT_CHOICE=�I�����Ă������� (0-5): "

if "%OPT_CHOICE%"=="1" call :TOGGLE_OPTION USE_REGEX
if "%OPT_CHOICE%"=="2" call :TOGGLE_OPTION CASE_SENSITIVE
if "%OPT_CHOICE%"=="3" call :TOGGLE_OPTION INCLUDE_SUBDIRS
if "%OPT_CHOICE%"=="4" call :TOGGLE_OPTION BACKUP_MODE
if "%OPT_CHOICE%"=="5" call :SET_LOG_FILE
if "%OPT_CHOICE%"=="0" goto :EOF

goto :SET_OPTIONS

:TOGGLE_OPTION
set "VAR_NAME=%1"
if "!!VAR_NAME!!"=="0" (
    set "%VAR_NAME%=1"
    echo %C_GREEN%!VAR_NAME! ��L���ɂ��܂���%C_RESET%
) else (
    set "%VAR_NAME%=0"
    echo %C_YELLOW%!VAR_NAME! �𖳌��ɂ��܂���%C_RESET%
)
pause
goto :EOF

:SET_LOG_FILE
echo.
set /p "NEW_LOG=���O�t�@�C���p�X����� (�󔒂Ŗ���): "
set "LOG_FILE=!NEW_LOG!"
if "!LOG_FILE!"=="" (
    echo %C_YELLOW%���O�o�͂𖳌��ɂ��܂���%C_RESET%
) else (
    echo %C_GREEN%���O�t�@�C��: !LOG_FILE!%C_RESET%
)
pause
goto :EOF

:PREVIEW_RENAME
if "!SEARCH_PATTERN!"=="" (
    echo %C_RED%�G���[: �����p�^�[�����ݒ肳��Ă��܂���%C_RESET%
    pause
    goto :EOF
)

echo.
echo %C_CYAN%�v���r���[���s��...%C_RESET%
echo.

set "RENAME_COUNT=0"
call :PROCESS_FILES 1

echo.
echo %C_GREEN%�v���r���[����: !RENAME_COUNT! ���̃t�@�C�������l�[���Ώۂł�%C_RESET%
pause
goto :EOF

:EXECUTE_RENAME
if "!SEARCH_PATTERN!"=="" (
    echo %C_RED%�G���[: �����p�^�[�����ݒ肳��Ă��܂���%C_RESET%
    pause
    goto :EOF
)

echo.
echo %C_YELLOW%�x��: ���ۂɃt�@�C�������l�[�����܂�%C_RESET%
set /p "CONFIRM=���s���܂����H (y/N): "
if /i not "!CONFIRM!"=="y" goto :EOF

echo.
echo %C_CYAN%���l�[�����s��...%C_RESET%
echo.

if not "!LOG_FILE!"=="" (
    echo [%DATE% %TIME%] ���l�[�������J�n > "!LOG_FILE!"
)

set "RENAME_COUNT=0"
set "ERROR_COUNT=0"
call :PROCESS_FILES 0

echo.
echo %C_GREEN%��������: !RENAME_COUNT! ���̃t�@�C�������l�[�����܂���%C_RESET%
if !ERROR_COUNT! gtr 0 echo %C_RED%�G���[: !ERROR_COUNT! ���̏����Ɏ��s���܂���%C_RESET%

if not "!LOG_FILE!"=="" (
    echo [%DATE% %TIME%] ��������: !RENAME_COUNT! ������, !ERROR_COUNT! �����s >> "!LOG_FILE!"
)

pause
goto :EOF

:PROCESS_FILES
set "IS_PREVIEW=%1"

if "!INCLUDE_SUBDIRS!"=="1" (
    set "SEARCH_OPTION=/s"
) else (
    set "SEARCH_OPTION="
)

pushd "!TARGET_DIR!"

for /f "delims=" %%F in ('dir /b !SEARCH_OPTION! "!FILE_FILTER!" 2^>nul') do (
    set "ORIGINAL_NAME=%%~nxF"
    set "NEW_NAME=!ORIGINAL_NAME!"
    
    :: �u������
    if "!USE_REGEX!"=="1" (
        :: ���K�\�����[�h�i�ȈՎ����j
        call :REGEX_REPLACE "!NEW_NAME!" "!SEARCH_PATTERN!" "!REPLACE_PATTERN!" NEW_NAME
    ) else (
        :: �ʏ�̕�����u��
        if "!CASE_SENSITIVE!"=="1" (
            set "NEW_NAME=!NEW_NAME:%%SEARCH_PATTERN%%=%%REPLACE_PATTERN%%!"
        ) else (
            call :CASE_INSENSITIVE_REPLACE "!NEW_NAME!" "!SEARCH_PATTERN!" "!REPLACE_PATTERN!" NEW_NAME
        )
    )
    
    :: �ύX�����邩�`�F�b�N
    if not "!ORIGINAL_NAME!"=="!NEW_NAME!" (
        set /a RENAME_COUNT+=1
        
        if "!IS_PREVIEW!"=="1" (
            echo %C_WHITE%!ORIGINAL_NAME!%C_RESET% %C_MAGENTA%��%C_RESET% %C_GREEN%!NEW_NAME!%C_RESET%
        ) else (
            :: �o�b�N�A�b�v�쐬
            if "!BACKUP_MODE!"=="1" (
                copy "!ORIGINAL_NAME!" "!ORIGINAL_NAME!.bak" >nul 2>&1
            )
            
            :: ���ۂ̃��l�[��
            ren "!ORIGINAL_NAME!" "!NEW_NAME!" 2>nul
            if errorlevel 1 (
                set /a ERROR_COUNT+=1
                echo %C_RED%�G���[: !ORIGINAL_NAME! �̃��l�[���Ɏ��s%C_RESET%
                if not "!LOG_FILE!"=="" (
                    echo [%DATE% %TIME%] �G���[: !ORIGINAL_NAME! �� !NEW_NAME! >> "!LOG_FILE!"
                )
            ) else (
                echo %C_WHITE%!ORIGINAL_NAME!%C_RESET% %C_MAGENTA%��%C_RESET% %C_GREEN%!NEW_NAME!%C_RESET%
                if not "!LOG_FILE!"=="" (
                    echo [%DATE% %TIME%] ����: !ORIGINAL_NAME! �� !NEW_NAME! >> "!LOG_FILE!"
                )
            )
        )
    )
)

popd
goto :EOF

:CASE_INSENSITIVE_REPLACE
set "STR=%~1"
set "SEARCH=%~2"
set "REPLACE=%~3"
set "RESULT_VAR=%~4"

:: �啶������������ʂ��Ȃ��u���i�ȈՎ����j
set "TEMP_STR=!STR!"
for %%A in (a b c d e f g h i j k l m n o p q r s t u v w x y z) do (
    for %%B in (A B C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
        set "TEMP_STR=!TEMP_STR:%%A=%%B!"
    )
)
set "TEMP_SEARCH=!SEARCH!"
for %%A in (a b c d e f g h i j k l m n o p q r s t u v w x y z) do (
    for %%B in (A B C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
        set "TEMP_SEARCH=!TEMP_SEARCH:%%A=%%B!"
    )
)

if not "!TEMP_STR!"=="!TEMP_STR:%TEMP_SEARCH%=%REPLACE%!" (
    set "%RESULT_VAR%=!STR:%SEARCH%=%REPLACE%!"
) else (
    set "%RESULT_VAR%=!STR!"
)
goto :EOF

:REGEX_REPLACE
:: �ȈՐ��K�\�������i��{�I�ȃp�^�[���̂݁j
set "STR=%~1"
set "PATTERN=%~2"
set "REPLACE=%~3"
set "RESULT_VAR=%~4"

:: �����ł͊�{�I�Ȓu���̂ݎ���
set "%RESULT_VAR%=!STR:%PATTERN%=%REPLACE%!"
goto :EOF

:SHOW_HELP
cls
call :SHOW_HEADER
echo.
echo %C_GREEN%�g�p���@:%C_RESET%
echo   %C_WHITE%�C���^���N�e�B�u���[�h:%C_RESET%
echo     %SCRIPT_NAME%.bat
echo.
echo   %C_WHITE%�R�}���h���C�����[�h:%C_RESET%
echo     %SCRIPT_NAME%.bat -d "�f�B���N�g��" -s "����" -r "�u��" [�I�v�V����]
echo.
echo %C_GREEN%�I�v�V����:%C_RESET%
echo   %C_WHITE%-d DIR%C_RESET%      �Ώۃf�B���N�g��
echo   %C_WHITE%-s PATTERN%C_RESET%  �����p�^�[��
echo   %C_WHITE%-r PATTERN%C_RESET%  �u���p�^�[��
echo   %C_WHITE%-f FILTER%C_RESET%   �t�@�C���t�B���^ (�f�t�H���g: *.*)
echo   %C_WHITE%-p%C_RESET%          �v���r���[���[�h�Ŏ��s
echo   %C_WHITE%-x%C_RESET%          ���K�\�����g�p
echo   %C_WHITE%-c%C_RESET%          �啶�������������
echo   %C_WHITE%-sub%C_RESET%        �T�u�f�B���N�g�����Ώ�
echo   %C_WHITE%-b%C_RESET%          �o�b�N�A�b�v���쐬
echo   %C_WHITE%-l FILE%C_RESET%     ���O�t�@�C�����w��
echo.
echo %C_GREEN%��:%C_RESET%
echo   %C_YELLOW%IMG �� Photo �ɒu��:%C_RESET%
echo     %SCRIPT_NAME%.bat -s "IMG" -r "Photo" -p
echo.
echo   %C_YELLOW%�g���q��ύX:%C_RESET%
echo     %SCRIPT_NAME%.bat -f "*.jpeg" -s ".jpeg" -r ".jpg"
echo.
pause
goto :EOF

:PARSE_ARGS
:ARG_LOOP
if "%~1"=="" goto :EOF
if "%~1"=="-d" (
    set "TARGET_DIR=%~2"
    shift & shift
    goto :ARG_LOOP
)
if "%~1"=="-s" (
    set "SEARCH_PATTERN=%~2"
    shift & shift
    goto :ARG_LOOP
)
if "%~1"=="-r" (
    set "REPLACE_PATTERN=%~2"
    shift & shift
    goto :ARG_LOOP
)
if "%~1"=="-f" (
    set "FILE_FILTER=%~2"
    shift & shift
    goto :ARG_LOOP
)
if "%~1"=="-l" (
    set "LOG_FILE=%~2"
    shift & shift
    goto :ARG_LOOP
)
if "%~1"=="-p" (
    set "PREVIEW_MODE=1"
    shift
    goto :ARG_LOOP
)
if "%~1"=="-x" (
    set "USE_REGEX=1"
    shift
    goto :ARG_LOOP
)
if "%~1"=="-c" (
    set "CASE_SENSITIVE=1"
    shift
    goto :ARG_LOOP
)
if "%~1"=="-sub" (
    set "INCLUDE_SUBDIRS=1"
    shift
    goto :ARG_LOOP
)
if "%~1"=="-b" (
    set "BACKUP_MODE=1"
    shift
    goto :ARG_LOOP
)
shift
goto :ARG_LOOP

:END
echo.
echo %C_CYAN%�����p���肪�Ƃ��������܂����B%C_RESET%
pause >nul
exit /b 0
