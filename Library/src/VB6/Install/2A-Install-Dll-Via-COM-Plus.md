# Install the XpressNet DLL using the COM+ bridge

The XpressNet DLL is a 32-bit COM component.  
If you build your own application as **64-bit**, the DLL cannot be loaded in-process.  
Windows can solve this by running the DLL in a **COM+ server application** (an out-of-process host).  
Your 64-bit application can then still use the DLL via normal COM calls.

---

## 1. Prerequisites
- Administrator rights on the PC.
- The DLL must already be copied to a permanent location, e.g.:  
  `C:\Program Files\XpressNet\XpressNet.dll`
- The DLL must be registered in the **32-bit COM registry**:  
  ```bat
  C:\Windows\SysWOW64\regsvr32.exe "C:\Program Files\XpressNet\XpressNet.dll"
  ```

---

## 2. Create a COM+ application

1. Open the Windows start menu, type **Component Services**, and start the tool.  
   (Alternatively run `dcomcnfg`.)
2. In the left tree, open:  
   `Component Services → Computers → My Computer → COM+ Applications`
3. Right-click **COM+ Applications** → **New → Application…**
4. In the wizard:  
   - Choose **Server application**  
   - Name it for example **XpressNetHost**  
   - For Identity choose **Interactive user** (simplest setting)

---

## 3. Import the DLL into the COM+ application

1. Right-click your new application (e.g. *XpressNetHost*) → **New → Component…**  
2. Select **Install new component** and browse to your DLL  
   (`C:\Program Files\XpressNet\XpressNet.dll`).  
3. Confirm the wizard. The classes from the DLL now appear under *Components*.

---

## 4. Using from your application

- You do **not** need to change your code.  
- Simply create the object as usual:  
  ```vbnet
  Dim xp As New XpressNet.XpressNetClass
  ```
- If your application is 64-bit, Windows will now start a 32-bit **dllhost.exe** process for the COM+ application.  
- All calls are automatically marshalled between your process and the COM+ host.

---

## 5. Removing the COM+ application

If you ever want to remove the COM+ bridge:

1. Open **Component Services** again.  
2. Find your application (e.g. *XpressNetHost*) under **COM+ Applications**.  
3. Right-click and choose **Delete**.  
4. If you also want to unregister the DLL:  
   ```bat
   C:\Windows\SysWOW64\regsvr32.exe /u "C:\Program Files\XpressNet\XpressNet.dll"
   ```

---

## 6. Troubleshooting

- If you see *Class not registered* errors:  
  → Run the `regsvr32` command again (step 1) with admin rights.  
- If calls hang or fail:  
  → Check the Windows **Event Viewer → Applications and Services Logs → Microsoft → Windows → COM+**.  
- If you want more control:  
  → You can adjust identity, security, or idle timeout in the COM+ application’s **Properties** dialog.

---

### Summary
- 32-bit applications (VB6, VB.NET x86) can use the DLL directly.  
- 64-bit applications cannot load the DLL directly.  
- With the COM+ bridge, a 64-bit application can still use the DLL transparently via COM.
