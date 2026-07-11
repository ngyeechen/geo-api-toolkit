Step 1: Ensure you have an internet connection.
Step 2: Open 'latlonpairs.xlsx' and input lat lon of your origin-destination pairs. Save and close.
Step 3: Within the same folder, double-click on 'getdurationasync_latlonpairs_getattr.exe'.

Each file in this folder is for the following purposes:

1) 'latlonpairs.xlsx' :  This file contains the lat lon of both origin and destination that you want to generate drivetime for. Open this file and input the desired lat lon pairs as needed and save the file upon exit.

2) 'getdurationasync_latlonpairs_getattr.exe' : This executable file runs upon double-clicking. As long as you place the 'latlonpairs.xlsx' within the same folder as this executable file, the run can be attempted. CMD (i.e., a black windowed terminal) will be displayed upon execution which tells you if the execution completed successfully or otherwise. Upon success, a 'latlonpairsWDuration.xlsx' file will appear within the same folder and the terminal will disappear. Upon failure, no file will be generated. In detail, this code pulls the data via API from OSRM.

3) Possible causes of failure:

3.1) Incorrect dir: as I've stated, ensure you place the .exe file within the same folder as the latlonpairs.xlsx first before you perform point (2) above.

3.2) 'latlonpairs.xlsx' format has been changed accidentally : See the screenshot of input file given if you can't revert the changes.

3.3) 'latlonpairsWDuration.xlsx' already exist : if this file exist before you run it, the terminal will tell you this "Resuming from existing output file..."

