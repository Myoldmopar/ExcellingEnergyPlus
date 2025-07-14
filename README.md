# Excelling EnergyPlus

OK, so I'm not the first person to try to call EnergyPlus from Excel.  Even though I was doing it back in 2008.  But I might be one of the first to integrate the EnergyPlus API with callbacks into the VBA?  Maybe?  Lots of complications getting this going, and it's definitely definitely not robust against all issues, but it is an end-to-end prototype.

## Basic Architecture

- This is meant to be a simple workbook that can call EnergyPlus and report simulation status live from the running simulation.
- EnergyPlus includes a C API layer that exports a bunch of functions.  The relevant ones for here are functions that:
  - Allow the client to register callback functions that EnergyPlus will call back into periodically
  - Allow the client to launch an EnergyPlus simulation
- This is a Windows-only project setup, for now.  Maybe there is a future where it's expanded to Mac, but it's just a side quest at the moment, so not now.
- This requires EnergyPlus to be installed separately
- The architecture of EnergyPlus, and the architecture of the trampoline asset downloaded here, should match the architecture of your running MS Office:
  - Right now the VBA code is hardwired to 32-bit architecture, but could be generalized for either 32 or 64
  - Also right now, the trampoline DLL that comes with the release assets here are also built targeting 32 bit to match
 
## OK, so did you say trampoline?

Yes.  The trampoline layer is a small DLL layer that supports marshaling data properly between VBA and EnergyPlus' C API.
When EnergyPlus is built, the C API functions are, by default, built with a __cdecl calling convention.
This works great for many clients, but not VBA!
For VBA, when it calls into a DLL and there are arguments, or callbacks, involved, it needs them as __stdcall.
To avoid having to rebuild EnergyPlus purely for stdcall, there is a trampoline DLL that can talk to both calling conventions, and bounces data back and forth.
The trampoline accepts __stdcall for the functions that get called from VBA, but then knows how to make __cdecl function calls out to EnergyPlus.
SILLY!

## Alright, so what should I do?

- Download EnergyPlus
  - Note that for some reason the 32-bit build of 25.1.0 is missing the python312.dll file, weird!
  - I had to install Python 3.12 and then copy the python312.dll file over into the EnergyPlus install
- Download the latest release package here, should be 32-bit for now
- Open the workbook from this release package
- Select paths accordingly - the path to EnergyPlus install folder, the IDF/EPW to run, and an output directory
- Run EnergyPlus, and hopefully watch the live messages and progress!
  
