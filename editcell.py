from scriptforge import ScriptForge, CreateScriptService
import os
import uno
import sys
import calc

def set_dlg_text(dlg):
    """
    ダイアログにアクティブセルのテキストを表示
    """
    text = dlg.Controls('TextField1')
    doc = CreateScriptService("SFDocuments.Calc")
    cell = calc.get_active_cell()
    if cell:
        text.Value = doc.GetValue(cell.AbsoluteName) 

def set_cell_text(dlg):
    """
    アクティブセルのテキストを設定
    """
    text = dlg.Controls('TextField1')
    doc = CreateScriptService("SFDocuments.Calc")
    cell = calc.get_active_cell()
    if cell:
        doc.SetValue(cell.AbsoluteName,text.Value) 

def exec_dialog(event:uno):
    dlg = CreateScriptService('SFDialogs.Dialog', 'GlobalScope', "Standard", "editcell_python")
    set_dlg_text(dlg)
    rc = dlg.Execute()
    if rc == dlg.OKBUTTON:
        set_cell_text(dlg)
    dlg.Terminate()

'''
イベントハンドラ
'''
def refresh(event:uno):
    """
    ダイアログにアクティブセルのテキストを再表示
    """
    button = CreateScriptService('SFDialogs.DialogEvent', event)
    set_dlg_text(button.Parent)

def update(event:uno):
    """セルに書き込む"""
    button = CreateScriptService('SFDialogs.DialogEvent', event)
    set_cell_text(button.Parent)

def up(event:uno):
    """↑ボタン"""
    button = CreateScriptService('SFDialogs.DialogEvent', event)
    set_cell_text(button.Parent)
    if calc.activate_cell_offset(0,-1):
        set_dlg_text(button.Parent)
    
def down(event:uno):
    """↓ボタン"""
    button = CreateScriptService('SFDialogs.DialogEvent', event)
    set_cell_text(button.Parent)
    if calc.activate_cell_offset(0,1):
        set_dlg_text(button.Parent)

if __name__ == '__main__':
    #プログラムとして実行された場合(LOの外から)
    try:
        ScriptForge(hostname='localhost', port=2021)    #リスニングモードのLOに接続
    except:
        print("LOへの接続エラー")
        sys.exit(0)
    
    #calcファイルを開く
    ui = CreateScriptService("UI")
    docs = ui.Documents()
    path=os.path.join(os.getcwd(), "editcell.ods")
    url = "file:///{}".format(path.replace("\\","/"))
    if not url in docs:   
        doc = ui.OpenDocument(path)
    
    exec_dialog(None)
