@echo off 
cd /d "C:\Users\KEUMHWAN\Desktop\claude code\업무자동화\" 
if not exist ".git" exit /b 0 
git add -A 
git diff --cached --quiet || ( git commit -m "auto %date% %time:~0,5%" --quiet & git push origin main --quiet ) 
