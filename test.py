# -*- coding: utf_8 -*-
import uno
import unohelper
from com.sun.star.awt import XActionListener    #@UnresolvedImport

class MyActionListener( unohelper.Base, XActionListener ):
    def __init__(self,callback):
        self._callback = callback

    def actionPerformed(self, actionEvent):
        self._callback()

#dialog class
class MyDialog():

    #disp text of active cell in dialog
    def __init__(self):
        #create dialog
        if IS_EXECUTED_EXTERNAL:
            ctx = XSCRIPTCONTEXT.getComponentContext()  #get context from XSCRIPTCONTEXT
            dp = ctx.getServiceManager().createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
        else:
            localctx = uno.getComponentContext()  #get local component context
            dp = localctx.getServiceManager().createInstanceWithContext("com.sun.star.awt.DialogProvider", localctx)
        
        #create dialog as member
        self.oDialog = dp.createDialog("vnd.sun.star.script:Standard.Dialog1?location=application")
        
        self.redisp()
    
        #set callback funcs
        self.oDialog.getControl("button_close").addActionListener(MyActionListener(self.end_dialog))
        self.oDialog.getControl("button_ok").addActionListener(MyActionListener(self.ok))
        self.oDialog.getControl("button_redisp").addActionListener(MyActionListener(self.redisp))
        self.oDialog.getControl("button_up").addActionListener(MyActionListener(self.up))
        self.oDialog.getControl("button_down").addActionListener(MyActionListener(self.down))
        self.oDialog.getControl("button_close").addActionListener(MyActionListener(self.end_dialog))
          
        #display dialog
        self.oDialog.execute()
        
    #redisp text of active cell in dialog
    def redisp(self):
        oText =self.oDialog.getControl ("TextField1")
        oCell = get_active_cell()
        if oCell:
            oText.Text = oCell.Text.String
    
    '''
    callback funcs
    '''
    
    #set text of dialog to active cell
    def set_str(self):
        oText =self.oDialog.getControl ("TextField1")
        oCell = get_active_cell()
        if oCell.Text.String!=oText.Text:
            oCell.Text.String=oText.Text
    
    def ok(self):
        self.set_str()
        self.oDialog.endExecute()
    
    #↑
    def up(self):
        self.set_str()
        if activate_cell_offset(0,-1):
            self.redisp()
    
    #↓
    def down(self):
        self.set_str()
        if activate_cell_offset(0,1):
            self.redisp()
    
    def end_dialog(self):
        self.oDialog.endExecute()        

def disp_str(arg):
    dialog = MyDialog()

 
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
