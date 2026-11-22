"""One command to rebuild the model: build -> scenario runs -> final recalculation."""
import subprocess, sys, os
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.abspath(os.path.join(HERE, "..", "BCatcher_Group_Integrated_Financial_Model.xlsx"))
RECALC = "/root/.claude/skills/synced/e3a10683-0f6f-4df7-9d97-1f3de7e0ea41_a12ffef0-b97c-4fce-adde-69bad4852925/xlsx/scripts/recalc.py"
run = lambda *a: subprocess.run([sys.executable, *a], check=True, cwd=HERE)
run("build_model.py", OUT)
run("run_scenarios.py", OUT)
run(RECALC, OUT, "240")
print("Model ready:", OUT)
