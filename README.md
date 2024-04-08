<h1>MS-Word - Colourize text based on rsidR values</h1>
<h6>
This script will prompt you for the following:<br><br>
1 - A DOCx input file that you wish to process and colourize the text in it based on revisions.<br><br>
2 - An output DOCx file (new file) which will be the colourized version of the input file.<br><br>
3 - A text file that contains one rsidR value per line. You can easily create one by running the script to parse MS Word documents found here: https://github.com/jjrboucher/MS-Word-Parser<br>
    Navigate to the RSID worksheet. Enable column filtering (format as a table in Excel, or via Data, Filter option).<br>
	Filter on rsid type for rsidR only. Filter on count, excluding any with a count of 0. No need to try and colourize text for rsidR values that are not present in document.xml, since that means no text was entered in that session.<br>
	Sort by the count column in descending order (so that the rsidR with the highest count is first).<br>
	Copy the RSID values, and paste into a text file and save it. This becomes the rsidR text file that you point to at this step in the process.<br>
	<h4>Commenting out RSIDs in the text file</h4>
 	You can put a pound sign "#" at the start of a line with an RSID to comment it out. This allows you to make more than one version of the file, each colourizing different text. This is also necessary if you have a large file with more than 30 RSID values in it.<br><br>
 
4 - A JSON file containing colour options for text colours. The current file in this repository has 39 colours. You can edit them, delete some, change the order (via the numerical key), or add new ones.<br>
    See https://www.rapidtables.com/web/color/RGB_Color.html to get the numerical codes for the different colours.<br>
	Make sure you do not skip any numbers. The script wil still work. But you may miss some of the colour options. If you have 30 rsidR values to apply a colour, and you enter 30 colours but skip #s 15 & 22 (meaning the last two colours will be 31 & 32), those last 
        two colours will be missed. The script loops over the RSID values in the RSID text file and increments a counter and looks for a colour option with that # in the colour file. Meaning it will never reach 31 (it will stop after looping through all 30 RSID values).<br>
	<br>
	The format must follow the same as you see in the included file. You must have a comma after each entry, and no comma after the last entry.<br>
	The Python script reads this file and will colourize the text associated to the RSIDs in the order of colour in this JSON file (again, numerical order, not necessarily the order in the RSID text file).<br><br>

<h2>Sample colour entries</h2>
Here is an example of a few colour entries. You must follow this format, with no comma after the very last colour in your list.<br>
	<br>"1": ["red1", [255, 0, 0]],
	<br>"2": ["orange1", [255, 128, 0]],
	<br>"3": ["yellow1", [255, 255, 0]],
	<br>"4": ["green2", [128, 255, 0]],

<h2>Logging</h2>
The script will create a .log file in addition to the colourized file. The .log file will bear the same name as the output file you give the script, with .log as an extension.

<h2>Dependencies</h2>

<h6>If running the script on a Linux system, you may need to install python-tk. You can do this with the following
command on a Debian (e.g. Ubuntu) system from the terminal window:<br>  
    
    sudo apt-get install python3-tk
<br>
Whether running on Linux, Mac, or Windows, you may need to install some of the libraries if they are not included in
your installation of Python 3.11.
<br>
In particular, you may need to install docx==0.2.4.  
    
<br>You can do so as follows from a terminal window while in the folder with the script and requirements.txt file:

    pip3 install -r requirements.txt
<hr>
If any other libraries are missing when trying to execute the script, install those in the same manner.</h6>

<h2>Executable Version</h2>
If you'd rather run the executable rather than needing Python, grab the .exe file.<br>
<br>

