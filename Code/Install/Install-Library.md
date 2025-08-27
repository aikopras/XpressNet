# Installation of the XpressNet Library


There are three ways to install and use the Library (source code / DLL file). Each way has its own advantages and disadvantages. A distinction should be made between the development machine, and the production machines.

## 1: Add source code to the application.
During the development and test phase, it is possible to add the source code of the library to the application source code. For this purpose, within the IDE, a VB group can be created, which merges two projects: 1) the application code and 2) the XpressNet library code. As long as work is done within the IDE, it is possible to develop and test the program, without the need for administrator privileges. However, once an executable (.exe) is desired, administrator privileges are required to register the DLL. In such case, VB6 should be started with administrator privileges (see below under option 2)

**Advantages:**</br>
- No administrator privileges are needed during testing within the IDE.</br>
- If necessary, the XpressNet library code can be modified relatively easily.

**Disadvantages:**</br>
- Without administrator privileges it is not possible to create an executable (.exe)</br>
- Application program ***must*** be written in VB6 (and can not be VB.NET).</br>
- Details of the library (the source code) are visible to the application developer. This could (accidentally) corrupt the library, causing the whole thing to fail.</br>

**Installation:** see this [manual](1-VB6_Group.md) for details.

## 2: Register in windows registry
The DLL file (XpressNet.dll) is installed in the windows registry. This should be done on the development but also the production machines.

**Advantages:**</br>
- Easy to use.</br>
- Details of the library (such as the source code) remain hidden from the application developer.</br>
- The library (dll file) needs to be installed (=registered) only once on each machine. After that, any application program (.exe file) can use it.</br>
- Also works with applications not written in VB6, for example VB.NET (VB2019/2022), C#, or other .NET languages. Note: in that case the application should be built as 32-bit (x86), or in case of 64-bit applications, a COM+ bridge must be used (see [separate instructions ](2A-Install-Dll-Via-COM-Plus.md)).

**Disadvantages:**</br>
- Installation of the DLL requires administrator rights, not only on the development machine, but also all production machines.</br>
- Different versions of the DLL can cause problems.</br>
- The DLL file must be stored in a directory that may not be changed later.

**Installation:** </br>
Installation is required, in addition to the development machine, for all production machines. There are two methods to install the DLL in the windows registry.</br>
1. By using ***regsvr32***. This method is relatively simple, but is sometimes discouraged on the Internet. See this [installation guide](2A-Install-Dll-Via-regsvr32.md) how this can be done.</br>
2. The library code is compiled by VB6 and then registered (invisibly to the user) by VB6. This method is a bit more cumbersome, but avoids using regsvr32. How to do this is explained in this [installation guide](2B-Install-Dll-Via-VB6.md). Note that this method can only be used with VB6, and not with newer versions.

## 3: Distribute the DLL along with the application.
A third possibility is to put the XpressNet DLL file alongside the application program (.exe file) on the production machines. In this case, the DLL file does *not* have to be installed on the production machines in the registry, so no administrator rights are required. This is called *RegFree* installation.

On the development machine, however, the DLL must be installed in the registry. For this purpose, one of the methods mentioned under 2. above can be used.

**Advantages:**</br>
- No administrator rights needed on production machines.
- Details of the library (such as the source code) remain hidden from the application developer.</br>

**Disadvantages:**</br>
- Administrator rights are required on the development machine.</br>
- An additional class (.cls) file must be added to the application source code, and one line in the application program must be modified.

**Installation:** </br>
The source code should be downloaded as described under Option 1. In addition an extra class file (XPNRegFree.cls) should be added. Details are described in the example application named [Regfree_test](../../Examples/Examples.md).
