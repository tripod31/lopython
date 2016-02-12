# -*- coding: utf_8 -*-
#import uno

def show_message(desktop, message):
    """shows message."""
    frame = desktop.getCurrentFrame()
    window = frame.getContainerWindow()
    toolkit = window.getToolkit()
    msgbox = toolkit.createMessageBox(
        window,"MESSAGEBOX", 1, 'massage', message)
    return msgbox.execute()


def get_active_cell():
    try:
        doc = XSCRIPTCONTEXT.getDocument()
        selected = doc.getCurrentSelection()
        if selected.supportsService('com.sun.star.sheet.SheetCellRange'):
            addr = selected.getRangeAddress()
            
            txt = 'Column: %s\nRow: %s' % (addr.EndColumn, addr.EndRow)
            show_message(XSCRIPTCONTEXT.getDesktop(), txt)
    except Exception as e:
        print(str(e))

   
if __name__ == '__main__':
    import uno 
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext )
    ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
    doc = desktop.loadComponentFromURL( "private:factory/scalc","_blank", 0, () )
    
    import unopy
    XSCRIPTCONTEXT = unopy.ScriptContext(ctx)
    
    get_active_cell()