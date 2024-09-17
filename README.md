# FileList2Csv
---
## set the paths and parameters
![set the paths](images/08.png))
## set the Task Schedule
1. open the Windows Task Schedule
2. select Create Basic Task and enter the task name.
![Task Schedule](images/01.png)
3. Task Trigger, select Daily.
![Task Trigger](images/02.png)
4. set the time to run.
![set the time](images/03.png)
5. Action, select Start a Program.
![Action](images/04.png)
6. Enter the Task parameters.
- Program = cscript
- Arguments = full path to the FileList2Csv.vbs file on your computer
- Start in = the same path in the arguments (only path , not include te filename)
![Start a Program](images/05.png))
![Task parameters](images/06.png))
7. Finish.
![Finish](images/06.png))