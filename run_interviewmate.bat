@echo off
title InterviewMate
echo =====================================
echo        INTERVIEWMATE - STARTING
echo =====================================
python -m pip install -r requirements.txt
streamlit run app.py
pause
