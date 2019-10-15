#pragma code_seg("PAGE")
bool SafeReadchar(char* cRetdata,LONG dwMemAddress)
{
bool retvar=false;
__try
	{
		*cRetdata=*(char*)(dwMemAddress);
	}
__except(1)
	{
	retvar=true;
	}

return !retvar;
}
#pragma code_seg("PAGE")
void ReadKernelMem(LONG dwAddr,LONG dwSize ,char* cReturnbytes)
{
char* Mem=(char*)ExAllocatePool(PagedPool,dwSize);
LONG dwExitSize;
bool IsNo1ReadSuccess=true;
for (int i=0;i<dwSize;i++)
{
	dwExitSize=i;
	if(!SafeReadchar((char*)((LONG)Mem+dwExitSize),dwAddr+dwExitSize))
	{
		if(i==0) 
		{
		IsNo1ReadSuccess=false; 
		}else
		{
		IsNo1ReadSuccess=true;
		}
	break;
	}
}
	if (IsNo1ReadSuccess)
	{
	*(PLONG)cReturnbytes=dwExitSize+1;
	}else
	{
	*(PLONG)cReturnbytes=0;
	}
RtlMoveMemory((char*)((LONG)cReturnbytes+4), Mem,dwExitSize+1);
dwExitSize=0;
ExFreePool(Mem);
Mem=NULL;
}
#pragma code_seg("PAGE")
LONG WriteKernelMem(LONG lpWriteAddr, char* lpWriteBytes,LONG lpWriteSize)
{
long retvar;
__try
{
	memmove((char*)lpWriteAddr,lpWriteBytes,lpWriteSize);
	retvar= STATUS_SUCCESS;
}__except(1)
{
	retvar= -1;
}
return retvar;
}