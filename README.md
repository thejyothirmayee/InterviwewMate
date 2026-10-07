# InterviewMate — Complete Placement Preparation App

InterviewMate is a local Streamlit workspace for students preparing for campus placements.

## Main modules
- Login / Create Account using email, phone number or username
- Resume validation and resume intelligence
- Company + exact job role matching
- Job-description validation
- Skill match and missing-skill analysis
- ATS/role-focused resume tailoring
- PDF/DOCX tailored resume export
- Aptitude Arena: Quantitative, Logical Reasoning, Verbal Ability
- Coding Lab: beginner Python placement problems + syntax/review + reference solutions
- Technical Round: Python, SQL/DBMS, AI/ML, Deep Learning/NLP/CV, CS Fundamentals
- Project Round: project-defense questions personalized from the resume
- HR Round: behavioral and recruiter questions
- Mock Interview: HR / Technical / Project simulation, one question at a time
- Personalized 7-day placement preparation roadmap
- Placement Readiness score and round-by-round progress dashboard
- SQLite storage for accounts and practice attempts

## Run on Windows
```powershell
python -m pip install -r requirements.txt
streamlit run app.py
```

## First launch
Open **Create Account** and make your own credentials. There are no fixed default credentials.

## Important
The Coding Lab intentionally does not execute arbitrary user-submitted code. It performs local Python syntax/review checks and provides a reference solution to keep the local student app safer.
