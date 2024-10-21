package main

import (
	"fmt"
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

var (
	modoleaut32 = windows.NewLazySystemDLL("oleaut32.dll")
	modole32    = windows.NewLazySystemDLL("ole32.dll")

	procCoInitialize     = modole32.NewProc("CoInitialize")
	procCoUninitialize   = modole32.NewProc("CoUninitialize")
	procCoCreateInstance = modole32.NewProc("CoCreateInstance")
)

const (
	CLSCTX_INPROC_SERVER = 1
	IID_IUnknown         = "{00000000-0000-0000-C000-000000000046}"
)

// IUnknown is the most basic COM interface.
type IUnknown struct {
	vtbl *IUnknownVtbl
}

type IUnknownVtbl struct {
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
}

func (u *IUnknown) Release() uint32 {
	ret, _, _ := syscall.Syscall(u.vtbl.Release, 1,
		uintptr(unsafe.Pointer(u)),
		0,
		0)
	return uint32(ret)
}

func main() {
	// Initialize COM
	procCoInitialize.Call(0)
	defer procCoUninitialize.Call()

	// Define the CLSID for UIAutomation
	clsid, err := windows.GUIDFromString("{ff48dba4-60ef-4201-aa87-54103eef594e}")
	if err != nil {
		panic(err)
	}

	// Define the IID for IUIAutomation
	iid, err := windows.GUIDFromString("{30cbe57d-d9d0-452a-ab13-7ac5ac4825ee}")
	if err != nil {
		panic(err)
	}

	var automation *IUnknown
	hr, _, _ := procCoCreateInstance.Call(
		uintptr(unsafe.Pointer(&clsid)),
		0,
		CLSCTX_INPROC_SERVER,
		uintptr(unsafe.Pointer(&iid)),
		uintptr(unsafe.Pointer(&automation)),
	)

	if hr != 0 {
		fmt.Printf("CoCreateInstance failed: %x\n", hr)
		return
	}

	fmt.Println("Successfully created UI Automation object")

	// Don't forget to release the COM object
	automation.Release()
}
