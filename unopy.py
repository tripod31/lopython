#
#!     # unopy.py
# -*- coding: utf_8 -*-
 
import uno
import unohelper
 
from com.sun.star.script.provider import XScriptContext #@UnresolvedImport
 
class ScriptContext(unohelper.Base, XScriptContext):
    def __init__(self, ctx):
        self.ctx = ctx
    
    def getComponentContext(self):
        return self.ctx
    
    def getDesktop(self):
        return self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
                
    def getDocument(self):
        return self.getDesktop().getCurrentComponent()
 
def connect():
    ctx = None
    try:
        localctx = uno.getComponentContext()
        resolver = localctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localctx)
        ctx = resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        if ctx:
            return ScriptContext(ctx)
    except:
        pass
    return None

def run_script(ctx, mod_name, func_name, location="user"):
    script_url = "vnd.sun.star.script:%s$%s?language=Python&location=%s" % (mod_name, func_name, location)
    msp = ctx.getValueByName("/singletons/com.sun.star.script.provider.theMasterScriptProviderFactory")
    sp = msp.createScriptProvider("")
    script = sp.getScript(script_url)
    return script.invoke((), (), ())
