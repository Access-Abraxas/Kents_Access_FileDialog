# Kents_Access_FileDialog
Kent Gorrell's project for using the MSO FileDialog tools for Microsoft Access 


## How to use these files in your Access database

1. In the Access database you want to use the dialog in, import in the Tables 
   and VBA Modules from the **accdb\Kents_FileDialog_Tools.accdb** database in 
   this project.

2. Then, in your VBA Code, where you need to show a File Dialog to get a string
   path to a file location, call the following VBA code: 

```
Dim strFilePath As String
strFilePath = Select_FullPath("Select Files", "", "Access", "*.accdb;*.accde", False)
```


## Source Code Overview
The source code for Kent's Access Dialog project is provided in this project 
in a very specific format.  Because Microsoft Access databases are binary files, 
a text diff of those files is not particularly useful for looking at changes in 
the Access database code.  So, for this project, along with the development 
Access database, I've created some tools to unzip Access Template files, so that
templates can be created from the Access development database and all of database 
objects and code for this project can be uploaded as text files for inclusion in 
the repository for this project.  This is useful because then is it easy to see 
the changes to the Access database in a text-based diff, each time a change is 
uploaded.


## Source Code Description
The following is a description for the files and directories included in this
project:


1. The "accdb" dir:

   * Kents_FileDialog_Tools.accdb 
     The development Access ACCDB database for this project.

   * Kents_FileDialog_Tools.accdt
     The Access ACCDT template file created from the development database. 

      
2. The "images" dir: 

   * icon.jpg -
     The icon file created for this project.

   * icon_sm.jpg -
     The small version of the icon created for this project.

   * Screenshot.jpg -
     The screenshot image used to create the ACCDT file.


3. The "scripts" dir: 

   * AssembleAccTemplate.bat -
     A batch file to build an Access ACCDT template file from the files
     currently in the "src" folder in this project.  This ACCDT file is 
     built it into the "bin" directory for this project.

   * UnzipAccTemplate.bat -
     A batch file to expand the Access ACCDT template file, from the accdt
     file in the "accdb" directory, into the "src" directory. This will 
     delete any existing files in the "src" directory.


4. The "src" dir:

   * The Access Template (ACCDT) Source Files - 
     The "src" directory contains all of the expanded files and directories 
     from the Access ACCDT template file, found in the "accdb" directory for 
     this project.  These file were created by running the 
     "UnzipAccTemplate.bat" script.  These files are all of the files needed
     to create an Access database from an Access template, which are in a text
     based format, so that we can track changes made to the original source 
     Access database found in the "accdb" folder of this project.


5. Files in the project root:

   * ChangeLog.txt - 
     A human description of each of the changes made to this project.  

   * README.md -    
     This initial README file created by GitHub for this project.

