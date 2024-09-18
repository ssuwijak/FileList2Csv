# FileList2Csv
### Description.
> this vbscript is for scanning all files & folders under the specified path. and write them down into a .csv file which can be opened by Excel.
### How to use the script in command line mode.
> `cscript FileLit2Csv.vbs "path\to\be\scanned" "path\of\the\csv\outout\file" scan_subdir_or_not "csv_separator"`
### How to use the script by Windows Task Scheduler.
1. Open the code and hardcode your paths and some parameters.
![set the paths](images/10.png))
2. open the Windows Task Schedule, select Create Basic Task and then enter the task name.
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
![Start a Program](images/11.png))
![Task parameters](images/12.png))
7. Finish.
![Finish](images/06.png))
