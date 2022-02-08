from scriptforge import ScriptForge, CreateScriptService
import os
import uno
import sys

'''
問題点
・アクティブセルが取れない
'''

def get_cell_text(dlg):
    """
    ダイアログにアクティブセルのテキストを表示
    """
    text = dlg.Controls('TextField1')
    doc = CreateScriptService("SFDocuments.Calc")
    addr = doc.LastCell("Sheet1")
    text.Value = doc.GetValue(addr) 

def set_cell_text(dlg):
    """
    アクティブセルのテキストを設定
    """
    text = dlg.Controls('TextField1')
    doc = CreateScriptService("SFDocuments.Calc")
    addr = doc.LastCell("Sheet1")
    doc.SetValue(addr,text.Value) 

def exec_dialog(event:uno):
    dlg = CreateScriptService('SFDialogs.Dialog', 'GlobalScope', "Standard", "Dialog2")
    get_cell_text(dlg)
    rc = dlg.Execute()
    if rc == dlg.OKBUTTON:
        set_cell_text(dlg)
    dlg.Terminate()

'''
イベント
'''
def redisp(event:uno):
    """
    ダイアログにアクティブセルのテキストを再表示
    """
    button = CreateScriptService('SFDialogs.DialogEvent', event)
    get_cell_text(button.Parent)

if __name__ == '__main__':
    #プログラムとして実行された場合(LOの外から)
    try:
        ScriptForge(hostname='localhost', port=2021)    #リスニングモードのLOに接続
    except:
        print("LOへの接続エラー")
        sys.exit(0)
    
    #test.odsを開く
    ui = CreateScriptService("UI")
    docs = ui.Documents()
    path=os.path.join(os.getcwd(), "test.ods")
    url = "file:///{}".format(path.replace("\\","/"))
    if not url in docs:   
        doc = ui.OpenDocument(path)

    exec_dialog(None)
    #doc.close(True)
