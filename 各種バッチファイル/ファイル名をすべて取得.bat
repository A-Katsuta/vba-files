@echo off

:: ���̃o�b�`�t�@�C��������t�H���_���̃t�@�C�����̈ꗗ���擾���܂��B
:: �T�u�t�H���_��o�b�`�t�@�C�����g�̖��O�͏��O���܂��B

echo �t�@�C�����̈ꗗ���쐬���Ă��܂�...

dir /b /a-d | findstr /v /i /c:"%~nx0" > �t�@�C�������X�g.txt

echo.
echo �������܂����B
echo �u�t�@�C�������X�g.txt�v�Ƃ������O�Ńt�@�C�����쐬����܂����B
echo ���̃t�@�C�����J���ē��e���R�s�[���AExcel�ɓ\��t���Ă��������B
echo.
pause