python macro sample for LibreOffice Calc
=====
It's eclipse project,simple python macro sample for LibreOffice Calc.  
It displays content of active cell in dialog.  
It can move active cell to up/downward.  
You can modify text in textbox.   

<img src="http://www.geocities.jp/tripod31hoge/images/lopython.jpg"/>

develpment environment
-----
###### Windows
+ Windows10Proffesional64bit
+ Eclipse 4.5.1  
+ Eclipse Pydev Plugin 4.5.4  
+ LibreOffce 5.1  

###### Linux
+ debian8
+ Eclipse 4.5.1
+ Eclipse Pydev Plugin 4.5.4
+ LibreOffce 4.3

Files
-----
+ Dialog1.xdl   
Dialog definition.

+ test.py  
macro module.  
Main function is "disp_str".

+ unopy.py  
This is needed when run test.py externaly,not macro.

+ test.ods  
Calc file with button to execute macro.

Setup LibreOffice
-----
#### Import Dialog1.xdl   
import this dialog to libre office.

#### copy test.py to LibreOffice directory
This is needed to run script as macro.  
In my case the directory is  

###### Windows
System directory  
```
D:\Program Files (x86)\LibreOffice 5\share\Scripts\python\
```
User directory is below,but when I execute script in the directory,error occurs.  
```
C:\Users\[user]\AppData\Roaming\LibreOffice\4\user\scripts\python\
```

###### Linux  
System directory   
```
/usr/lib/libreoffice/share/Scripts/python/
```
Or User Directory  
```
~/.config/libreoffice/4/user/Scripts/
```

Setup Eclipse
-----
This is needed to run script in Eclipse.  
#### Add python interpreter
This is needed in Windows Only.  

###### Windows
Set path of python bandled with LO.  
Set it in "Window"->"preference".  
In my case the path is  
```
D:\Program Files (x86)\LibreOffice 5\program\python.exe
```
#### Select interpreter to project
In project property->"Pydev Interpreter/Grammer".  
###### Windows
Set Interpreter above.
###### Linux
Set python3 interpreter.

#### Add library path
In project property->"Pydev PYTHONPATH"->"External Libraries"  

###### Windows
In my case the path is  
```
D:\Program Files (x86)\LibreOffice 5\program
```

###### Linux
Install "python3-uno" package.Librariy path is added automatically.

To run script in Eclipse
-----
#### Lunch LibreOffice in listening mode
This is needed before you run script in Eclise.  
I use Ant task to do this.  
Edit LibreOfiice executable path in build.xml.  
###### Windows
In my case the path is  
```
D:\Program Files (x86)\LibreOffice 5\program\soffice.exe
```  
In Eclipse,select build.xml->"Run As"->"Ant Build..."->"Taget"Tab->Check "exec_soffice"=>"Run".

###### Linux
In Eclipse,select build.xml->"Run As"->"Ant Build..."->"Taget"Tab->Check "exec_soffice_linux"=>"Run".

Known problems
-----
###### I can't run script in user directory
When I run the scirpt in user directory from LO,error occurs.  
In case of linux,I can run the script.  

###### BreakPoint in callback functions of dialog,does'nt work
When I debug the script in Eclipse,it doesn't stop at breakpoint,in callback functions.
