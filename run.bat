@echo off
cd /d %~dp0
set STREAMLIT_LOG_LEVEL=error
streamlit run main.py --logger.level=error
pause