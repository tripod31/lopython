from scriptforge import ScriptForge, CreateScriptService
import os
import uno
import sys

'''
メモ：
Scriptforge.BasicのThisComponentで、アクティブなドキュメントのオブジェクトを取得できる。マクロとしてでなく、LO外部から実行した場合も
'''
def activate_cell_offset(argOffsetColumn,argOffsetRow):
    '''
    :returns:    True:Success False:Fail
    '''    
    oActiveCell = get_active_cell()
    if oActiveCell is None:
        return False
    else:
        new_cell = get_cell_offset(oActiveCell, argOffsetColumn, argOffsetRow)
        if new_cell:
            activate_cell(new_cell)
        else:
            return False
        
        return True
    
def get_cell_offset(oCell,argOffsetColumn,argOffsetRow):
    if oCell.ImplementationName=="ScCellObj":
        lngStartColumn = oCell.CellAddress.Column + argOffsetColumn
        lngStartRow = oCell.CellAddress.Row + argOffsetRow
        oSheet = oCell.Spreadsheet
        try:
            retCell = oSheet.getCellByPosition(lngStartColumn, lngStartRow)
        except:
            retCell = None
        return retCell
    else:
        return None

def activate_cell(oCell):
    if oCell.ImplementationName=='ScCellObj':
        svc = CreateScriptService("Basic")
        doc = svc.ThisComponent
        oController = doc.getCurrentController()
        oController.setActiveSheet(oCell.Spreadsheet)
        oController.select(oCell)

def get_active_cell():
    '''
    returns:
        scCellObj
    '''
    svc = CreateScriptService("Basic")
    doc = svc.ThisComponent
    sel = doc.CurrentSelection
    if sel.ImplementationName == "ScCellObj":
        return sel
    else:
        return None

def get_cell_text(dlg):
    """
    ダイアログにアクティブセルのテキストを表示
    """
    text = dlg.Controls('TextField1')
    doc = CreateScriptService("SFDocuments.Calc")
    cell = get_active_cell()
    if cell:
        text.Value = doc.GetValue(cell.AbsoluteName) 

def set_cell_text(dlg):
    """
    アクティブセルのテキストを設定
    """
    text = dlg.Controls('TextField1')
    doc = CreateScriptService("SFDocuments.Calc")
    cell = get_active_cell()
    if cell:
        doc.SetValue(cell.AbsoluteName,text.Value) 

def exec_dialog(event:uno):
    dlg = CreateScriptService('SFDialogs.Dialog', 'GlobalScope', "Standard", "Dialog2")
    get_cell_text(dlg)
    rc = dlg.Execute()
    if rc == dlg.OKBUTTON:
        set_cell_text(dlg)
    dlg.Terminate()

'''
イベントハンドラ
'''
def redisp(event:uno):
    """
    ダイアログにアクティブセルのテキストを再表示
    """
    button = CreateScriptService('SFDialogs.DialogEvent', event)
    get_cell_text(button.Parent)

def up(event:uno):
    """↑ボタン"""
    button = CreateScriptService('SFDialogs.DialogEvent', event)
    set_cell_text(button.Parent)
    if activate_cell_offset(0,-1):
        get_cell_text(button.Parent)
    
def down(event:uno):
    """↓"""
    button = CreateScriptService('SFDialogs.DialogEvent', event)
    set_cell_text(button.Parent)
    if activate_cell_offset(0,1):
        get_cell_text(button.Parent)

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
    path=os.path.join(os.getcwd(), "test_sf.ods")
    url = "file:///{}".format(path.replace("\\","/"))
    if not url in docs:   
        doc = ui.OpenDocument(path)
    
    exec_dialog(None)
