# VB6ARC
Name given to collection of some old programs written by me in Visual Basic 6. 
The programs code were written around year 2010-12. 

Since these tools were written when I was in my 9-10th standard, they don't follow any standard coding practise and even documentation is missing. I am adding whatever I can recall, you might need to recompile from source to make them usable again.

## The Collection

### [Process Killer](Process_Killer)
A very usefull tool to restore your computer in a certain state - You can simply take a snapshot of processes and with a click of button you can kill any new processes created after the snapshot. It can also prevent any new processes from creating. 
You can read more about it [here](Process_Killer).

### [USB Manager](USB_Manager)
Prevent annoying people from plugging in any device, it can hide plugged devices !

### [Drive Manager](Drive_Manager)
It can hide and show drives, even disable them just with a click. You don't need to go in `Administrative Tools > Computer Management > Disk Managment > Assign/Remove Drive letter`, to hide a drive !

### [HTTP Server](HTTP_Server)
A standalone server which can only server static files. The server doesn't has any additional requirement.
Note: winsock could be missing in newer system, I am not sure.

### [Instant Shutdown](Instant_ShutDown)
Program which can shutdown your machine without asking and waiting for anything ! Best in situation when you just want to shutdown it !

### [Local RDP](Local_RDP)
A small server based on HTTP Server, It could stream your computer's screen over network, which can be accessed over any standard browser.

# Directory structure

- `build` each project has its own build directory which contains an old compiled executable.
- `src` contains source code of the compiled binary. 

 
