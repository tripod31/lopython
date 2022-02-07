# pylint: disable=import-error

import uno
import unohelper
from com.sun.star.awt import XActionListener
import os

class MyActionListener( unohelper.Base, XActionListener ):
    def __init__(self,callback):
        self._callback = callback

    def actionPerformed(self, actionEvent):
        self._callback()


class MyDialog():
    """dialog class"""

    def __init__(self):
        """disp text of active cell in dialog"""
        
        '''
        create dialog.We have to use different way to get context,whether it is invoked as macro,or executed as script.
        '''
        if IS_EXECUTED_EXTERNAL:
            #when it is executed as script
            ctx = XSCRIPTCONTEXT.getComponentContext()  #get context from XSCRIPTCONTEXT
            dp = ctx.getServiceManager().createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
        else:
            #when it is invoked as macro
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
        
    
    def redisp(self):
        """re-display text of active cell in dialog"""
        
        oText =self.oDialog.getControl ("TextField1")
        oCell = get_active_cell()
        if oCell:
            oText.Text = oCell.Text.String
    
    '''
    callback funcs
    '''
    
    def set_str(self):
        oText =self.oDialog.getControl ("TextField1")
        oCell = get_active_cell()
        if oCell.Text.String!=oText.Text:
            oCell.Text.String=oText.Text
    
    def ok(self):
        self.set_str()
        self.oDialog.endExecute()
    
    
    def up(self):
        """↑"""
        self.set_str()
        if activate_cell_offset(0,-1):
            self.redisp()
    
    def down(self):
        """↓"""
        self.set_str()
        if activate_cell_offset(0,1):
            self.redisp()
    
    def end_dialog(self):
        self.oDialog.endExecute()        


def disp_str(*dummy):
    '''
    main function
    
    :param:*dummy: dummy argument.to prevent argument number error when it is executed.
        Number of arguments are varied by how it is invoked,from macro menu,form button.
    '''
    _dialog = MyDialog()

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
    if not doc:
        return None
    selection = doc.getCurrentSelection()
    if selection.ImplementationName=='ScCellObj':
        return selection
    else:
        return None        
   
if __name__ == '__main__':
#    import uno
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext )
    ctx = resolver.resolve( "uno:socket,host=localhost,port=2021;urp;StarOffice.ComponentContext" )
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
    #doc = desktop.loadComponentFromURL( "private:factory/scalc","_blank", 0, () )   #open new calc document    
    #open test.ods
    path=os.path.join(os.getcwd(), "test.ods")
    doc = desktop.loadComponentFromURL( "file:///"+path,"_default", 0, () )   #open new calc document
    
    import unopy
    XSCRIPTCONTEXT = unopy.ScriptContext(ctx)
    IS_EXECUTED_EXTERNAL = True     #when excetuted as executable
    
    disp_str()
    #doc.close(True)
    
else:
    IS_EXECUTED_EXTERNAL = False    #When Executed as macro
