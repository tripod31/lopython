# pythonのLibreOffice Calc用マクロのサンプル
pythonのLibreOffice Calc用マクロのサンプルです。  
アクティブセルの内容をダイアログに表示します。  
テキストをダイアログで編集できます。  
## 動作確認環境
+ Windows11
+ vscode1.64.0
+ LibreOffce 7.2  
## ファイル
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
## LibreOffice設定
#### Dialog1.xdlをインポート   
#### test.pyをLibreOfficeディレクトリにコピー  
test.pyをマクロとして実行するために必要  
私の環境ではディレクトリは：  
```
C:\Users\(USER)\AppData\Roaming\LibreOffice\4\user\scripts\python\
```
## VSCode設定
VSCodeからスプリプトをデバッグするために必要  
#### pythonインタプリタを追加  
"Python:Default Interpreter Path"を設定
LibreOfficeに付属するpythonを登録    
私の環境でのパス：  
```
C:\Program Files\LibreOffice\program\python.exe
```
## VSCodeからスクリプトをデバッグ
#### LibreOfficeをリスニングモードで起動
スクリプトを実行する前に必要  
soffice.exeに引数を付けて起動  
```
"C:/Program Files/LibreOffice/program/soffice.exe" --accept='socket,host=localhost,port=2021;urp;'
```  
## 既知の問題
VSCodeでスクリプトをデバッグする時、ダイアログのコールバック関数内のブレークポイントで止まらない  
## ダイアログ画面
<img src="https://user-images.githubusercontent.com/6335693/50577268-4182f580-0e68-11e9-9c6e-129ac94d96d6.jpg"/>