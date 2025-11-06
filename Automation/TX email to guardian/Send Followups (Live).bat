@echo off
set "SCRIPT_PATH=C:\GoogleSync\Automation\TX email to guardian\send_followups_picker.py"
py -3 -c "import importlib.util, pathlib; p=pathlib.Path(r'%SCRIPT_PATH%'); spec=importlib.util.spec_from_file_location('followups', str(p)); mod=importlib.util.module_from_spec(spec); spec.loader.exec_module(mod); mod.DRY_RUN=False; mod.main()"
if errorlevel 1 pause
