extern "C"
{
#include <ntddk.h>
}
#include "ReadAndWriteMem.h"

typedef struct _DRIVER_INFO {
	LIST_ENTRY ListLink;
	LIST_ENTRY Link1;
	LIST_ENTRY InitializationLink;
	PVOID Base;
	PVOID Entry;
	ULONG Size;
	UNICODE_STRING PathName;
	UNICODE_STRING FileName;
}DRIVER_INFO,*PDRIVER_INFO;

#define READMEM CTL_CODE(FILE_DEVICE_UNKNOWN,0x800,METHOD_BUFFERED,FILE_ANY_ACCESS)
#define WRITEMEM CTL_CODE(FILE_DEVICE_UNKNOWN,0x801,METHOD_BUFFERED,FILE_ANY_ACCESS)

typedef struct _DEVZICE_EXTENSION {
	PDEVICE_OBJECT pDevice;
	UNICODE_STRING ustrDeviceName;
	UNICODE_STRING ustrSymLinkName;
} DEVICE_EXTENSION, *PDEVICE_EXTENSION;
#pragma code_seg("PAGE")
NTSTATUS MjFunction(PDEVICE_OBJECT pDevObject, PIRP pIrp)
{
	PIO_STACK_LOCATION pStack=IoGetCurrentIrpStackLocation(pIrp);
	ULONG InBuffersize=pStack->Parameters.DeviceIoControl.InputBufferLength;
	ULONG OutBuffersize=pStack->Parameters.DeviceIoControl.OutputBufferLength;
	ULONG Code=pStack->Parameters.DeviceIoControl.IoControlCode;

	switch(pStack->MajorFunction) {
	case IRP_MJ_CREATE:
	{break;}
	case IRP_MJ_CLOSE:
	{break;}
	case IRP_MJ_DEVICE_CONTROL:
	{
	PULONG sAddress=(PULONG)pIrp->AssociatedIrp.SystemBuffer;
		if (Code==READMEM)
		{//ReadMemory
		ReadKernelMem(*sAddress ,*(PULONG)((LONG)sAddress+4),(char*)sAddress);
		}
		if (Code==WRITEMEM)
		{
		*sAddress=WriteKernelMem(*sAddress,(char*)((LONG)sAddress+8),*(PLONG)((LONG)sAddress+4));
		}
	break;
	}
	}
	pIrp->IoStatus.Status=STATUS_SUCCESS;
	pIrp->IoStatus.Information=OutBuffersize;
	IoCompleteRequest(pIrp,IO_NO_INCREMENT);
	return STATUS_SUCCESS;
}
#pragma code_seg("PAGE")
void UnloadMyDriver(PDRIVER_OBJECT pDriverObject)
{
PDEVICE_OBJECT deviceObject = pDriverObject->DeviceObject;
UNICODE_STRING symbolicLinkName;

    RtlInitUnicodeString(&symbolicLinkName, L"\\??\\HanfSys_Disastrously");
    IoDeleteSymbolicLink(&symbolicLinkName);
    IoDeleteDevice(deviceObject);

    return;
}
#pragma code_seg("INIT")
extern "C" 
NTSTATUS 
DriverEntry(PDRIVER_OBJECT DriverObject,PUNICODE_STRING RegistryPath)
{
PDEVICE_OBJECT pDevice;
UNICODE_STRING ntDeviceName;
UNICODE_STRING symbolicLinkName;
PDEVICE_EXTENSION pDevExt;
NTSTATUS status;

RtlInitUnicodeString(&ntDeviceName, L"\\Device\\HanfDisastrously");

    status = IoCreateDevice(DriverObject,               // DriverObject
                            sizeof(DEVICE_EXTENSION), // DeviceExtensionSize
                            &ntDeviceName,              // DeviceName
                            FILE_DEVICE_UNKNOWN,        // DeviceType
                            0,    // DeviceCharacteristics
                            TRUE,                      // Not Exclusive
                            &pDevice               // DeviceObject
                           );

    if (!NT_SUCCESS(status))
	 {
        KdPrint(("IoCreateDevice returned 0x%x\n", status));
        return(status);
    }

  
	//创建符号链接 \??\sss'
     RtlInitUnicodeString(&symbolicLinkName, L"\\??\\HanfSys_Disastrously");
     status = IoCreateSymbolicLink(&symbolicLinkName, &ntDeviceName);
	 
	//得到设备扩展
	pDevExt = (PDEVICE_EXTENSION)pDevice->DeviceExtension;

	//设置扩展设备的设备对象
	pDevExt->pDevice = pDevice;

	//设置扩展设备的设备名称
    pDevExt->ustrDeviceName = ntDeviceName;
	pDevExt->ustrSymLinkName=symbolicLinkName;
	DriverObject->MajorFunction[IRP_MJ_CREATE]=MjFunction;
	DriverObject->MajorFunction[IRP_MJ_CLOSE]=MjFunction;
	DriverObject->MajorFunction[IRP_MJ_DEVICE_CONTROL]=MjFunction;

	DriverObject->DriverUnload = UnloadMyDriver;
	
     if (!NT_SUCCESS(status)) 
	{
         IoDeleteDevice(pDevice);
         //KdPrint(("IoCreateSymbolicLink returned 0x%x\n", status));
         return(status);
     }
//表示为IO设备
pDevice->Flags |= DO_BUFFERED_IO;
return status;
}