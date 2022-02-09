from scriptforge import CreateScriptService

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
