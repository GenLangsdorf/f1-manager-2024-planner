# Double-click this file (needs Python installed). It asks which save to use,
# updates that career's planner workbook, then opens it. Close the workbook in Excel first.
import os, subprocess, sys
here = os.path.dirname(os.path.abspath(__file__))
rc = subprocess.call([sys.executable, os.path.join(here, "f1m_update.py"), "--open"])
if rc:
    input("\nSomething went wrong (see above). Press Enter to close.")
