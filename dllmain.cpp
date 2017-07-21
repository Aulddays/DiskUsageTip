/****************************** Module Header ******************************\
Module Name:  dllmain.cpp
Project:      DiskUsageTip
Copyright (c) Aulddays.
Copyright (c) Microsoft Corporation.

The file implements DllMain, and the DllGetClassObject, DllCanUnloadNow,
DllRegisterServer, DllUnregisterServer functions that are necessary for a COM
DLL.

DllGetClassObject invokes the class factory defined in ClassFactory.h/cpp and
queries to the specific interface.

DllCanUnloadNow checks if we can unload the component from the memory.

DllRegisterServer registers the COM server and the context menu handler in
the registry by invoking the helper functions defined in Reg.h/cpp. The
context menu handler is associated with the .cpp file class.

DllUnregisterServer unregisters the COM server and the context menu handler.

install:
Regsvr32.exe DiskUsageTip.dll
uninstall:
Regsvr32.exe /u DiskUsageTip.dll


This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#include <windows.h>
#include <Guiddef.h>
#include "ClassFactory.h"           // For the class factory
#include "Reg.h"


// {7D586193-A8F7-4D86-B6A9-90BDF61413C2}
// When you write your own handler, you must create a new CLSID by using the 
// "Create GUID" tool in the Tools menu, and specify the CLSID value here.
const CLSID CLSID_FileContextMenuExt =
{ 0x7d586193, 0xa8f7, 0x4d86, { 0xb6, 0xa9, 0x90, 0xbd, 0xf6, 0x14, 0x13, 0xc2 } };


HINSTANCE   g_hInst = NULL;
long        g_cDllRef = 0;


BOOL APIENTRY DllMain(HMODULE hModule, DWORD dwReason, LPVOID lpReserved)
{
	switch (dwReason)
	{
	case DLL_PROCESS_ATTACH:
		// Hold the instance of this DLL module, we will use it to get the 
		// path of the DLL to register the component.
		g_hInst = hModule;
		DisableThreadLibraryCalls(hModule);
		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}


//
//   FUNCTION: DllGetClassObject
//
//   PURPOSE: Create the class factory and query to the specific interface.
//
//   PARAMETERS:
//   * rclsid - The CLSID that will associate the correct data and code.
//   * riid - A reference to the identifier of the interface that the caller 
//     is to use to communicate with the class object.
//   * ppv - The address of a pointer variable that receives the interface 
//     pointer requested in riid. Upon successful return, *ppv contains the 
//     requested interface pointer. If an error occurs, the interface pointer 
//     is NULL. 
//
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
	HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;

	if (IsEqualCLSID(CLSID_FileContextMenuExt, rclsid))
	{
		hr = E_OUTOFMEMORY;

		ClassFactory *pClassFactory = new ClassFactory();
		if (pClassFactory)
		{
			hr = pClassFactory->QueryInterface(riid, ppv);
			pClassFactory->Release();
		}
	}

	return hr;
}


//
//   FUNCTION: DllCanUnloadNow
//
//   PURPOSE: Check if we can unload the component from the memory.
//
//   NOTE: The component can be unloaded from the memory when its reference 
//   count is zero (i.e. nobody is still using the component).
// 
STDAPI DllCanUnloadNow(void)
{
	return g_cDllRef > 0 ? S_FALSE : S_OK;
}


//
//   FUNCTION: DllRegisterServer
//
//   PURPOSE: Register the COM server and the context menu handler.
// 
STDAPI DllRegisterServer(void)
{
	HRESULT hr;

	wchar_t szModule[MAX_PATH];
	if (GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule)) == 0)
	{
		hr = HRESULT_FROM_WIN32(GetLastError());
		return hr;
	}

	// Register the component.
	hr = RegisterInprocServer(szModule, CLSID_FileContextMenuExt, L"DiskUsageTip", L"Apartment");
	if (SUCCEEDED(hr))
	{
		// Register the context menu handler. The context menu handler is 
		// associated with the .cpp file class.
		hr = RegisterShellExtContextMenuHandler(L"Directory"/*L".cpp"*/, CLSID_FileContextMenuExt, L"DiskUsageTip");
		if (SUCCEEDED(hr))
			hr = RegisterShellExtContextMenuHandler(L"Drive", CLSID_FileContextMenuExt, L"DiskUsageTip");
	}

	return hr;
}


//
//   FUNCTION: DllUnregisterServer
//
//   PURPOSE: Unregister the COM server and the context menu handler.
// 
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = S_OK;

	wchar_t szModule[MAX_PATH];
	if (GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule)) == 0)
	{
		hr = HRESULT_FROM_WIN32(GetLastError());
		return hr;
	}

	// Unregister the component.
	hr = UnregisterInprocServer(CLSID_FileContextMenuExt);
	if (SUCCEEDED(hr))
	{
		// Unregister the context menu handler.
		hr = UnregisterShellExtContextMenuHandler(L"Directory"/*L".cpp"*/, CLSID_FileContextMenuExt);
		if (SUCCEEDED(hr))
			hr = UnregisterShellExtContextMenuHandler(L"Drive", CLSID_FileContextMenuExt);
	}

	return hr;
}