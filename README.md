<div align="center">

## Generic Multi\-Threading Engine  \(Asynchronous Drop in Replacement for VB's CallByName with Events\)


</div>

### Description



----

LAST UPDATED 16-Sep-04 11:50am 

----

This is an Asynchronous ActiveX server to replace the built in VB6 CallByName function. By substituting current CallByname instances with a call to the AsyncServer, this allows the user interface to remain responsive while heavy processes are being conducted in the background.

It also means that your application has effectively become multi-threaded.

Events within the AsyncServer notify the client app when the call has been completed. Events within the called DLL can also be raised as per normal (for interface notification etc..)

Hope it serves some purpose for someone. I use it across a range of projects in instances where it is important to maintain a responsive UI.

The other nice thing is that it is generic and existing classes/objects that you've created can be used without change.

''' *** Notes: Before using asyncronously, the AsyncServer.exe and AsyncTest.dll projects should be compiled. The AsyncTest DLL is purely included as an example of what the AsyncServer can be used for. The small test project (Project1) demonstrates the difference between a function in the test dll being called thro' the server or thro' the exe's thread.

''' ***

BTW - was going to include callbyname example but thought that interested parties would already have a handle on this.

the original zip has been updated to include a 'preprocessing' event that allows factory created objects to be tweaked before async processing continues.

Added implemented interface to AsyncServer to allow for VTable 'fast' tight integration.

changed the ReturnAsyncObject function parameters to accept an object in order to be identical to the CallByName implementation. Cahnged the FactoryObject event nad callbacks to be more meaningfully named PreProcessFactoryObject.
 
### More Info
 
Before using asyncronously, the AsyncServer project should be compiled. The AsyncTest DLL is purely included as an example of what the AsyncServer can be used for. The small test project (Project1) demonstrates the difference between a function in the test dll being called thro' the server or thro' the exe's thread.


<span>             |<span>
---                |---
**Submitted On**   |2004-09-16 12:03:34
**By**             |[jim tollan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jim-tollan.md)
**Level**          |Advanced
**User Rating**    |5.0 (50 globes from 10 users)
**Compatibility**  |VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Generic\_Mu1794389162004\.zip](https://github.com/Planet-Source-Code/jim-tollan-generic-multi-threading-engine-asynchronous-drop-in-replacement-for-vb-s-callby__1-56175/archive/master.zip)








