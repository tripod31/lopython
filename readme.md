pythonのLibreOffice Calc用マクロのサンプル
=====
pythonのLibreOffice Calc用マクロのサンプルの、Eclipse用プロジェクトです。  
アクティブセルの内容をダイアログに表示します。  
テキストをダイアログで編集できます。  

<img src="http://www.geocities.jp/tripod31hoge/images/lopython.jpg"/>

開発環境
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

ファイル
-----
+ Dialog1.xdl   
ダイアログ定義  

+ test.py  
マクロモジュール  
メインの関数は"disp_str"  

+ unopy.py  
test.pyをマクロとしてではなく、直接実行するために必要  

+ test.ods  
マクロを実行するCalcサンプル  

LibreOffice設定
-----
#### Dialog1.xdlをインポート   

#### test.pyをLibreOfficeディレクトリにコピー  
test.pyをマクロとして実行するために必要  
私の環境ではディレクトリは：  

###### Windows
システムディレクトリ  
```
D:\Program Files (x86)\LibreOffice 5\share\Scripts\python\
```
ユーザーディレクトリに置く場合は↓、しかしマクロを実行するとエラーになった    
```
C:\Users\[user]\AppData\Roaming\LibreOffice\4\user\scripts\python\
```

###### Linux  
システムディレクトリ   
```
/usr/lib/libreoffice/share/Scripts/python/
```
ユーザーディレクトリ  
```
~/.config/libreoffice/4/user/Scripts/
```

Eclipse設定
-----
Eclipseからスプリプトを実行するために必要  
#### pythonインタプリタを追加
Windowsの場合のみ必要    

###### Windows
LibreOfficeに付属するpythonを登録  
"Window"->"preference"で設定  
私の環境でのパス：  
```
D:\Program Files (x86)\LibreOffice 5\program\python.exe
```
#### インタプリタをプロジェクトで選択
プロジェクトのプロパティ->"Pydev Interpreter/Grammer".  
###### Windows
上記のインタプリタを設定
###### Linux
python3インタプリタを設定

#### ライブラリパスを追加
プロジェクトのプロパティ->"Pydev PYTHONPATH"->"External Libraries"  

###### Windows
私の環境でのパス：  
```
D:\Program Files (x86)\LibreOffice 5\program
```

###### Linux
"python3-uno"パッケージをインストールする。ライブラリパスは自動で設定される

Eclipseからスクリプトを実行
-----
#### LibreOfficeをリスニングモードで起動
私はAntタスクで起動している  
build.xmlの中のLibreOfficeのexeファイルのパスを編集する  
###### Windows
私の環境でのパス：  
```
D:\Program Files (x86)\LibreOffice 5\program\soffice.exe
```  
Eclipseで、build.xmlを選択->"Run As"->"Ant Build..."->"Taget"Tab->"exec_soffice"をクリック=>"Run".

###### Linux
Eclipseで、build.xmlを選択->"Run As"->"Ant Build..."->"Taget"Tab->"exec_soffice_linux"をクリック=>"Run".

既知の問題
-----
###### ユーザーディレクトリに置いたマクロをLibreOfficeから実行するとエラー
Windowsの場合のみ    

###### ダイアログのコールバック関数内のブレークポイントが動作しない
Eclipseでスクリプトをデバッグする時、コールバック関数内のブレークポイントで止まらない  
