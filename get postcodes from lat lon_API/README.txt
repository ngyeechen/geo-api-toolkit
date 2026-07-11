Step 1: Ensure you have an internet connection.
Step 2: Open 'input.xlsx' and input lat lon. Save and close.
Step 3: Within the same folder, double-click on 'getpostcode_input_getattr.exe'.

Each file in this folder is for the following purposes:

1) 'input.xlsx' :  This file contains the lat lon that you want to generate Postcodes for. Open this file and input the desired lat lon as needed and save the file upon exit.

2) 'getpostcode_input_getattr.exe' : This executable file runs upon double-clicking. As long as you place the 'input.xlsx' within the same folder as this executable file, the run can be attempted. CMD (i.e., a black windowed terminal) will be displayed upon execution which tells you if the execution completed successfully or otherwise. Upon success, a 'output_with_uk_postcodes.xlsx' file will appear within the same folder and the terminal will disappear. Upon failure, no file will be generated. In detail, this code pulls the data via API from Nominatim.

3) Possible causes of failure:

3.1) Incorrect dir : as I've stated, ensure you place the .exe file within the same folder as the input.xlsx first before you perform point (2) above.

3.2) 'input.xlsx' format has been changed accidentally : See the screenshot of input file given if you can't revert the changes.

3.3) 'output_with_uk_postcodes.xlsx' already exist : if this file exist before you run it, the terminal will tell you this "Resuming from existing output file..."


