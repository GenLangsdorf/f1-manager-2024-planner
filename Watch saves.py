# Double-click this file and leave the window open while you play. The planner is rebuilt
# every time the game saves, as long as the workbook is closed in Excel at that moment.
import os, subprocess, sys
here = os.path.dirname(os.path.abspath(__file__))
subprocess.call([sys.executable, os.path.join(here, "f1m_update.py"), "--watch"])
input("\nStopped. Press Enter to close.")
