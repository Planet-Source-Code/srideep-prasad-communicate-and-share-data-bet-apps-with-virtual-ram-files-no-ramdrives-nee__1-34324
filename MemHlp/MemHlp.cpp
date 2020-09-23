#include <windows.h>
#include <winuser.h>
#include <winbase.h>
void _stdcall ReadMem(char *Dest,int Addr)
{
    strcpy(Dest,(char *)Addr);
}


void _stdcall WriteMem(char *bData,int Addr)
{
    strcpy((char *)Addr,bData);
}

