<h1>MS-Word - Colourize text based on rsidR values</h1>

<h2>Overview</h2>
This script allows you to select a MS Word document. The script will parse through it and identify all rsidR values it finds and display them in a window to the left of the menu. You can click on the associated square to select the colour you wish to assign to all text that was typed in that session.

The rsidR values will be listed in decreasing order, with the one that appears the most often listed at the top of the list.

<h2>Saving a colour version</h2>
The script allows you to save a copy of the document with the colours applied to it. It will create a corresponding log file as well that lists the rsidR values that were coloured, and the RBG code for the colours that were used.

You can clear the colours applied to the on screen document and colour other rsidR values and save that to yet another document. This allows you to create as many versions as you want, each with select rsidR values coloured according to your selection. In each case, a corresponding log file will be created.

<h2>rsidR Report</h2>
The script also provides you with the option to create a report which will contain each rsidR value and a corresponding table. The table will have three columns. The page number, the paragraph number, and the associated text. This can be useful to produce first to identify text of interest and see the associated rsidR values. You can apply colour to those rsidR values rather than trial and error trying to find which rsidR value corresponds to some text of interest within the document.

<h2>Logging</h2>
The script will create a .log file in addition to the colourized file. The .log file will bear the same name as the output file you give the script, with .log as an extension.

<h2>Dependencies</h2>

python-docx>=1.1.0
pywin32>=306

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

