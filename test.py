# -*- coding: utf_8 -*-
import uno
import unohelper
from com.sun.star.awt import XActionListener    #@UnresolvedImport

class MyActionListener( unohelper.Base, XActionListener ):
    def __init__(self, dialog,callback):
        self.dialog = dialog
        self._callback = callback

    def actionPerformed(self, actionEvent):
        self._callback(self.dialog)

#disp text of active cell in dialog
def disp_str(arg):
    #create dialog
    if IS_EXECUTED_EXTERNAL:
        ctx = XSCRIPTCONTEXT.getComponentContext()  #get context from XSCRIPTCONTEXT
        dp = ctx.getServiceManager().createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    else:
        localctx = uno.getComponentContext()  #get local component context
        dp = localctx.getServiceManager().createInstanceWithContext("com.sun.star.awt.DialogProvider", localctx)
    
    oDialog = dp.createDialog("vnd.sun.star.script:Standard.Dialog1?location=application")
    
    redisp(oDialog)

    #set callback funcs
    oDialog.getControl("button_close").addActionListener(MyActionListener(oDialog,end_dialog))
    oDialog.getControl("button_ok").addActionListener(MyActionListener(oDialog,ok))
    oDialog.getControl("button_redisp").addActionListener(MyActionListener(oDialog,redisp))
    oDialog.getControl("button_up").addActionListener(MyActionListener(oDialog,up))
    oDialog.getControl("button_down").addActionListener(MyActionListener(oDialog,down))
    oDialog.getControl("button_close").addActionListener(MyActionListener(oDialog,end_dialog))
    
    down(oDialog)    
    #display dailog
    oDialog.execute()
    
#redisp text of active cell in dialog
def redisp(oDialog):
    oText =oDialog.getControl ("TextField1")
    oCell = get_active_cell()
    if oCell:
        oText.Text = oCell.Text.String

'''
callback funcs
'''

#set text of dialog to active cell
def set_str(oDialog):
    oText =oDialog.getControl ("TextField1")
    oCell = get_active_cell()
    if oCell.Text.String!=oText.Text:
        oCell.Text.String=oText.Text

def ok(oDialog):
    set_str(oDialog)
    oDialog.endExecute()

#↑
def up(oDialog):
    set_str(oDialog)
    if activate_cell_offset(0,-1):
        redisp(oDialog)

#↓
def down(oDialog):
    set_str(oDialog)
    if activate_cell_offset(0,1):
        redisp(oDialog)

def end_dialog(oDialog):
    oDialog.endExecute()        

 
'''
common funcs
'''
def show_message(desktop, message):
    """shows message."""
    frame = desktop.getCurrentFrame()
    window = frame.getContainerWindow()
    toolkit = window.getToolkit()
    msgbox = toolkit.createMessageBox(
        window,"MESSAGEBOX", 1, 'massage', message)
    return msgbox.execute()


'''
:returns:    True:Success False:Fail
'''
def activate_cell_offset(argOffsetColumn,argOffsetRow):
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
    
def get_cell_offset(argBaseRange,argOffsetColumn,argOffsetRow):
    if argBaseRange.ImplementationName=="ScCellObj":
        lngStartColumn = argBaseRange.CellAddress.Column + argOffsetColumn
        lngStartRow = argBaseRange.CellAddress.Row + argOffsetRow
        oSheet = argBaseRange.Spreadsheet
        try:
            oRange = oSheet.getCellByPosition(lngStartColumn, lngStartRow)
        except:
            oRange = None
        return oRange
    else:
        return None

def activate_cell(selection):
    if selection.ImplementationName=='ScCellObj':
        oController = XSCRIPTCONTEXT.getDocument().getCurrentController()
        oController.setActiveSheet(selection.Spreadsheet)
        oController.select(selection)
    
def get_active_cell():
    doc = XSCRIPTCONTEXT.getDocument()
    selection = doc.getCurrentSelection()
    if selection.ImplementationName=='ScCellObj':
        return selection
    else:
        return None        
   
if __name__ == '__main__':
#    import uno
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext )
    ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
    doc = desktop.loadComponentFromURL( "private:factory/scalc","_blank", 0, () )
    
    import unopy
    XSCRIPTCONTEXT = unopy.ScriptContext(ctx)
    IS_EXECUTED_EXTERNAL = True
    
    disp_str(None)
else:
    IS_EXECUTED_EXTERNAL = False
