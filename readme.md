python macro sample for LibreOffice Calc
=====
It's eclipse project,simple python macro sample for LibreOffice Calc.  
It displays content of active cell in dialog.  
It can move active cell to up/downward.  
You can modify text in textbox.   

<img src="http://www.geocities.jp/tripod31hoge/images/lopython.jpg"/>

develpment environment
-----
+ Eclipse 4.5.1  
+ Eclipse Pydev Plugin 4.5.4  
+ LibreOffce 5.1  

Files
-----
+ Dialog1.xdl   
Dialog difinition.

+ test.py  
macro module.  
Main function is "disp_str()".

+ unopy.py  
This is needed when run test.py externaly,not macro.

+ test.ods  
Calc file with button to execute macro.

Setup LibreOffice
-----
###### Import Dialog1.xdl   
import this dialog to libre office.

###### copy test.py to LibreOffice directory
This is needed to run script as macro.  
In my case the directory is  
```
D:\Program Files (x86)\LibreOffice 5\share\Scripts\python\
```

Setup Eclipse
-----
This is needed to run script in Eclipse.  
###### Add python interpreter
In Window->preference.  

In my case the path is  
```
D:\Program Files (x86)\LibreOffice 5\program\python.exe
```

###### Apply interpreter
In project property->"Pydev Interpreter/Grammer".  
Set Interpreter above.

###### Add library path
In project property->"Pydev PYTHONPATH"->"External Libraries"  
In my case the path is  
```
D:\Program Files (x86)\LibreOffice 5\program
```

To run script in Eclipse
-----
###### Lunch LibreOffice in listening mode
This is needed before you run script in Eclise.  
I use Ant task to do this.  
Edit LibreOfiice executable path in build.xml.  
In my case the path is  
```
D:\Program Files (x86)\LibreOffice 5\program\soffice.exe
```  
In Eclipse,select build.xml->"Run As"->"Ant Build..."->"Taget"Tab->Check "Exec soffice"=>"Run".

Known problems
-----
###### I can't run script in user directory
I put the script in user directory.  
```
C:\Users\[user]\AppData\Roaming\LibreOffice\4\user\scripts\python\
```
When I run the scirpt from LO,error occurs.

###### BreakPoint in callback functions of dialog,does'nt work
When I debug the script in Eclipse,it doesn't stop at breakpoint,in callback functions.
