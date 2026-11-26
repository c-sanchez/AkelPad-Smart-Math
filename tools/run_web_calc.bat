@echo off
if not exist "_site" goto NeedsBuild

echo Web Calc will be available here: http://127.0.0.1:8000
python -m http.server --directory _site
goto Done

:NeedsBuild
echo There is no "_site" directory. Run "build_web_calc.py" first.

:Done
