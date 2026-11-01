# Version history

## Version 1.3 (October 2026)

### XpressNet DLL (Code/VB6_Source)

**Corrections**

- `XpressNet-Class-Main.cls`
  - PoM commands: the low byte of the CV number was calculated as `CV Mod 255` instead of `CV Mod 256`. As a result CV 256 became CV 1, and CV 257 to 1024 were one too high.
  - PoM bit write (`CmdPomBitWrite`): `0xE8`, as in the V3.6 specification (3.2.28.3). Version 1.2 used `0x7C` from the V4 specification (3.43.3), which seems to be an error in that specification.
  - The version of the command station (`ZentraleVersion`) is compared as a number (`Val`). Before, it was compared as a string, or converted implicitly. A string comparison fails for versions of 10.0 and higher, and the implicit conversion depends on the Windows regional settings (decimal separator).
  - `CmdZ21GetExtAccessory` sent one data byte too many (Z21 specification 5.5).
  - Lok information: the speed byte is decoded according to the speed steps. Before, it was always decoded as 27/28 speed steps, which gave a wrong speed for 14 and 128 speed steps.
  - The Z21 firmware version is BCD coded (Z21 specification 2.15). `0x01 0x23` is now shown as 1.23 (was 1.35).
- `XpressNet-Class-Receive-Z21.cls`
  - The XpressNet message and its XOR byte are taken from the Z21 message that is being decoded. Before, they were taken from the input buffer, which was only correct if a UDP packet contained a single Z21 message.
  - A message with a XOR error is no longer decoded (it is still logged).
  - DataLen consists of two bytes (little endian). A DataLen below 4 no longer results in an endless loop, and a Z21 message longer than 23 bytes no longer results in "subscript out of range".
- `XpressNet-Class-Receive.cls`: if the input buffer is full, the complete XpressNet messages are processed first. Before, the remaining bytes were dropped and the buffer stayed full, so nothing could be received anymore.
- `XpressNet-Class-AccAuto.cls`: `emptyFifo` (called after a response time-out) also stops the deactivation timer. Before, the timer kept running with an empty FIFO, which resulted in an overflow error and a deactivation of a wrong address.

**Changes**

- `XpressNet-Class-Send.cls`: the response time-out (5 seconds, 90 seconds in programming mode) is used for TCP, USB and serial connections. Before, it was used for TCP only, so a missing response blocked the FIFO for USB and serial connections. The USB interface has the same rule as the LAN interface (LI-USB 1.1, Interface 23151 1.3). The keep alive message after a time-out is sent for TCP only.
- `XpressNet-Class-Send.cls`: no flow control for the Z21 (UDP). The Z21 does not answer every command (Z21 specification 1.1), so messages for the Z21 are now sent immediately. Before, the FIFO, including the keep alive message, could block until another XpressNet message from the Z21 arrived.
- `XpressNet-Class-AccAuto.cls`: after a deactivation there is a pause of `AccessoryInterval` / 2 (62 ms) before the next accessory is activated, as in the example of the Z21 specification (5.2.1). This limits the current if many accessories are switched.
- `XpressNet.vbp`: version 1.3 ("Version October 2026").

**Upgrade note**

The DLL project uses "No Compatibility". Version 1.3 therefore has new GUIDs. Applications that were compiled against an earlier version of the DLL must set the reference to the XpressNet DLL again (Project → References), and must be compiled again. For Visual Studio (.NET) applications, the COM reference must be added again.

### Examples

- `1-Basic_test`, `2-Extended_test`, `4_Multiple_Instance_test`, `5_Regfree_test`, `6-Project_Group`: `FunctionsZustandF29F68` stored the function groups 6 to 10 in `FB(4)` and `FB(5)`; they are now stored in `FB(6)` to `FB(10)`.
- Z21 forms (`3_Z21_Basic test`, `4_Multiple_Instance_test`, `5_Regfree_test`, `6-Project_Group`): the Steps list contains 14, 27, 28 and 128, and "Number of Z21Messages" is updated (new Timer control `TimerMessageCounter`).
- `5_Regfree_test`, Z21 form: `=>` and `<=` switched the wrong position, and Feedback requested RS-Bus feedback instead of the switch information.
- `6-Project_Group`: the DLL source files are updated to version 1.3 (they were from before August 2025).
- All executables are compiled again with DLL version 1.3. The project files now contain the reference to the DLL.

### Documentation

- `Examples.md` / `Voorbeelden.md`: description of the Basic test added; the executables of the Regfree test are in the subdirectory `Exe`; note for the VS2019 example about the COM reference.
- `2A-Install-Dll-Via-regsvr32.md` / `2A-Installeren-Dll-Via-regsvr32.md`: on 64-bit Windows (such as Windows 11) Step 9B applies, also for VB6 applications.
- `Version_History.md` / `Versiegeschiedenis.md` added.

## Version 1.2 (August 2025)

- The application program can set the directory for the logfiles (`LogDirectory`). See `1-Basic_test` for an example.
- The FIFO for XpressNet messages was enlarged from 32 to 128 entries.
- The time between two accessory commands (`AccessoryInterval`) was changed from 250 to 125 ms.

## Version 1.1 (December 2024)

- `CmdSwitchInfoRequest` was added, to request the switch information. `CmdFeedbackRequest` is now intended for (RS-Bus) feedback decoders.
- The examples request the correct feedback data, use the default Lenz and Roco IP addresses, and position 1 is now straight.

## Version 1.0 (November 2024)

- First version.
