#https://github.com/broli/PKB-Open-jobs-report
#https://github.com/broli/PKB-Open-jobs-report
# Main.py
# from openJobs_class import OpenJobsApp # If you kept the old filename
from app_shell import OpenJobsApp # If you renamed to app_shell.py

if __name__ == "__main__":
    app = OpenJobsApp()
    app.mainloop()