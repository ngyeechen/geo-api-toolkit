Step 1: Ensure you have an internet connection.
Step 2: Open 'input2.xlsx' and input postcodes. Save and close.
Step 3: Within the same folder, double-click on 'getlatlon_input2_getattr.exe'.

Each file in this folder is for the following purposes:

1) 'input2.xlsx' :  This file contains the Postcodes that you want to generate lat and lon for. Open this file and input the desired Postcodes as needed and save the file upon exit.

2) 'getlatlon_input2_getattr.exe' : This executable file runs upon double-clicking. As long as you place the 'input2.xlsx' within the same folder as this executable file, the run can be attempted. CMD (i.e., a black windowed terminal) will be displayed upon execution which tells you if the execution completed successfully or otherwise. Upon success, a 'output_with_lat_lon.xlsx' file will appear within the same folder and the terminal will disappear. Upon failure, no file will be generated. In detail, this code pulls the data via API from Postcodes.io → findthatpostcode.uk → Nominatim, specifically in this order.

3) Possible causes of failure:

3.1) Incorrect dir : as I've stated, ensure you place the .exe file within the same folder as the input2.xlsx first before you perform point (2) above.

3.2) 'input2.xlsx' format has been changed accidentally : See the screenshot of input file given if you can't revert the changes.

3.3) 'output_with_lat_lon.xlsx' already exist : if this file exist before you run it, the terminal will tell you this "Resuming from existing output file..."

4) postcode_cache.pkl : just a cache/"memory" that stores all the postcodes that have been run previously such that your current run will take much shorter time. Try removing this cache and then perform point (2) as stated above and you'd find it'll take longer to produce the result that's all.

