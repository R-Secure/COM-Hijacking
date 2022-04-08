### COM-Hijacking

Some of the default Windows scheduled tasks use triggers to call COM objects. This is very useful for gaining persistence, as we can look for scheduled tasks with specific triggers. This program looks for the ```At log on of any user``` trigger which will provide you with reboot-persistence. To leverage this, the program creates the required entries in HKCU pointing to the payload you specified as a command line argument (see usage) to have it loaded every time a user logs in.

### Usage:

```
COMHijack.exe C:\Windows\Temp\beacon.dll
```
