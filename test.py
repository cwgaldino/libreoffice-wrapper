import uno
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
context = resolver.resolve("uno:socket,host=localhost,port=8100;urp;StarOffice.ComponentContext")

desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
document = desktop.getCurrentComponent()

cursor = document.Text.createTextCursor()
document.Text.insertString(cursor, "This text is being added to openoffice using python and uno package.", 0)
