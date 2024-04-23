@ECHO OFF
::Run in case of virtual environment
CALL %userprofile%\rwsbn\Scripts\activate
::Correct absolute path to code file
python %userprofile%\rwsbn\main.py

PAUSE