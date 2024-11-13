# Installation of the XpressNet Library


There are three ways to install and use the Library (source code / dll file). Each way has its own advantages and disadvantages. A distinction should be made between the development machine, and the production machines.

## 1: Add source code to the application.
It is possible to add the source code of the library to the application source code. For this purpose a VB group can be created, which merges two projects: 1) the application code and 2) the XpressNet library code.

**Advantages:**</br>
- No administrator rights are required anywhere.</br>
- No DLL file is needed (after all, we are using the XpressNet source code). So no problems with registering DLLs.</br>
- If necessary, the XpressNet library code can be modified relatively easily.</br>
- After the application program is developed and an .exe file is created, only this .exe file needs to be put on the production machines.

**Disadvantages:**</br>
- Application program must be written in VB6.</br>
- Details of the library (the source code) are visible to the application developer. This could (accidentally) corrupt the library, causing the whole thing to fail.</br>

**Installation:** see this [manual](1-VB6_Group.md) for details.

## 2: Register in windows registry
The dll file (XpressNet.dll) is installed in the windows registry. This should be done on the development but also the production machines.

**Advantages:**</br>
- Easy to use.</br>
- Details of the library (such as the source code) remain hidden from the application developer.</br>
- The library (dll file) needs to be installed (=registered) only once on each machine. After that, any application program (.exe file) can use it.</br>
- Should also work with applications not written in VB6, but in C and .Net, for example.

**Disadvantages:**</br>
- Installation of the DLL requires administrator rights, not only on the development machine, but also all production machines.</br>
- Different versions of the dll can cause problems.

**Installation:** </br>
Installation is required, in addition to the development machine, for all production machines. There are two methods to install the DLL in the windows registry.</br>
1. By using ***regsvr32***. This method is relatively simple, but is sometimes discouraged on the Internet. See this [installation guide](2A-Install-Dll-Via-regsvr32.md) how this can be done.</br>
2. The library code is compiled by VB6 and then registered (invisibly to the user) by VB6. This method is a bit more cumbersome, but avoid using regsvr32. How to do this is explained in this [installation guide](2B-Install-Dll-Via-VB6.md).


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
