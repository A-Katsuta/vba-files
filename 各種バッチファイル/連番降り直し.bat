@echo off
setlocal enabledelayedexpansion

:: =============================================================
:: �t�@�C���A�ԐU�蒼���o�b�`
:: �T�v: �����̔ԍ��𖳎����A�t�@�C���𖼑O���ɕ��ׂĘA�Ԃ�U�蒼��
:: =============================================================

:: --- �ݒ荀�� (���������ɍ��킹�ĕύX���Ă�������) ---

REM �ΏۂƂ���t�@�C�����̃p�^�[�����w�肵�܂� (���C���h�J�[�h * ���g���܂�)
set "FILE_PATTERN=�t�H�[�}�b�g*.txt"

REM �V�����t�@�C�����̐ړ���(�擪����)���w�肵�܂�
set "NEW_PREFIX=�t�H�[�}�b�g_������_"

:: ---------------------------------------------------------

echo �ȉ��̐ݒ�Ńt�@�C�����̕ύX�������J�n���܂��B
echo.
echo   �Ώۃt�@�C��: %FILE_PATTERN%
echo   �V�������O�@: %NEW_PREFIX%[�A��].�g���q
echo.
pause

set "count=0"

echo.
echo --- �����J�n ---

rem /on �Ńt�@�C���𖼑O���Ƀ\�[�g���Ă��珈������
for /f "delims=" %%F in ('dir /b /on "%FILE_PATTERN%"') do (
    rem �A�Ԃ�1���₷
    set /a count+=1

    rem 3���̃[�����ߔԍ����쐬 (��: 1 -> 001, 12 -> 012)
    set "num=000!count!"
    set "num=!num:~-3!"

    rem �V�����t�@�C�����𐶐� (�ړ��� + �[�����ߔԍ� + ���̊g���q)
    set "newName=!NEW_PREFIX!!num!%%~xF"

    rem �t�@�C������ύX
    ren "%%F" "!newName!"

    rem �������ʂ���ʂɕ\��
    echo [�ύX�O]: %%F  ==^>  [�ύX��]: !newName!
)

echo.
echo --- �������� ---
echo %count%�̃t�@�C�����������܂����B
pause