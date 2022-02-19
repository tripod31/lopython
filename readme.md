# LibreOffice Calc用pythonマクロのサンプル
アクティブセルの内容をダイアログに表示します。  
テキストをダイアログで編集できます。  
## 動作確認環境
+ Windows11
+ vscode1.64.0
+ LibreOffce 7.2  
## ファイル
-----
+ editcell_python.xdl   
ダイアログ定義  
+ editcell.py  
マクロモジュール
+ calc.py  
マクロモジュールから参照されるモジュール  
+ editcell.ods  
マクロを実行するCalcサンプル  
## LibreOffice設定
#### ダイアログ定義editcell_python.xdlをインポート   
#### editcell.pyをLibreOfficeディレクトリにコピー  
editcell.pyをマクロとして実行するために必要  
私の環境ではディレクトリは：  
```
C:\Users\(USER)\AppData\Roaming\LibreOffice\4\user\scripts\python\
```
#### calc.pyをLibreOfficeディレクトリにコピー  
マクロモジュールから別のモジュールを参照するために必要  
私の環境ではディレクトリは：  
```
C:\Users\(USER)\AppData\Roaming\LibreOffice\4\user\scripts\python\pythonpath\
```
pythonpathディレクトリを作成する必要がありました。  
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
"C:\Program Files\LibreOffice\program\soffice.exe" --accept='socket,host=localhost,port=2021;urp;'
```  
## 既知の問題
VSCodeでスクリプトをデバッグする時、ダイアログのイベントハンドラ関数内のブレークポイントで止まらない  
## ダイアログ画面
![editcell](https://user-images.githubusercontent.com/6335693/153326258-db64a78c-891e-4cf2-a36a-a54ddfd5dd6a.png)
