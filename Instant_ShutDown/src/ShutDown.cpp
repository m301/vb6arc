#define WINVER 0x0400
#include <windows.h> 
int main()
{
	HANDLE hToken;
	if (OpenProcessToken(GetCurrentProcess(), 
	TOKEN_QUERY|TOKEN_ADJUST_PRIVILEGES, &hToken))
	{
		TOKEN_PRIVILEGES tkp;

		LookupPrivilegeValue(NULL, SE_SHUTDOWN_NAME, &tkp.Privileges[0].Luid);

		tkp.PrivilegeCount = 1;
		tkp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED; 

		AdjustTokenPrivileges(hToken, FALSE, &tkp, 0, NULL, 0);
		CloseHandle(hToken);
	} // get system privileges. 

		/* get handle to dll */ 
   HINSTANCE hGetProcIDDLL = LoadLibrary("ntdll.dll"); 

   /* get pointer to the function in the dll*/ 
   FARPROC lpfnGetProcessID = GetProcAddress(HMODULE (hGetProcIDDLL),"NtShutdownSystem"); 

   /* 
	  Define the Function in the DLL for reuse. This is just prototyping the dll's function. 
	  A mock of it. Use "stdcall" for maximum compatibility. 
   */ 
   typedef int (__stdcall * pICFUNC)(DWORD); 

   pICFUNC shutitdown; 
   shutitdown = pICFUNC(lpfnGetProcessID); 

   /* The actual call to the function contained in the dll */ 
   shutitdown( 2 ); 

   /* Release the Dll */ 
   FreeLibrary(hGetProcIDDLL); 
}
