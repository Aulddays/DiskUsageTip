/****************************** Module Header ******************************\
Module Name:  DiskUsageTipExt.cpp
Project:      DiskUsageTip
Copyright (c) Aulddays.
Copyright (c) Microsoft Corporation.

Implementation of DiskUsageTipExt Shell context menu handler.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND,
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#define _CRT_SECURE_NO_WARNINGS
#define NOMINMAX	// disable the min/max macros

#include "DiskUsageTipExt.h"
#include "resource.h"
#include <strsafe.h>
#include <Shlwapi.h>
#pragma comment(lib, "shlwapi.lib")

#include <string>
#include <stdio.h>
#include <algorithm>

extern HINSTANCE g_hInst;
extern long g_cDllRef;

#define IDM_DETAIL             0  // The command's identifier offset

DiskUsageTipExt::DiskUsageTipExt(void) : m_cRef(1),
m_pszMenuText(L"%s of %s free (%0.2f%%)"),
m_pszVerb("diskusage"),
m_pwszVerb(L"diskusage"),
m_pszVerbCanonicalName("DiskUsageTip"),
m_pwszVerbCanonicalName(L"DiskUsageTip"),
m_pszVerbHelpText("Display Disk Usage"),
m_pwszVerbHelpText(L"Display Disk Usage")
{
	InterlockedIncrement(&g_cDllRef);

	m_diskUsageTip.resize(wcslen(m_pszMenuText) + 8 * 4 + 1);

	// Load the bitmap for the menu item. 
	// If you want the menu item bitmap to be transparent, the color depth of 
	// the bitmap must not be greater than 8bpp.
	//m_hMenuBmp = LoadImage(g_hInst, MAKEINTRESOURCE(IDB_OK),
	//	IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE | LR_LOADTRANSPARENT);

	HICON hIcon = (HICON)LoadImage(g_hInst, MAKEINTRESOURCE(IDI_DISK), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);
	int dstx = GetSystemMetrics(SM_CXMENUCHECK);
	int dsty = GetSystemMetrics(SM_CYMENUCHECK);
	dstx = std::max(dstx, 16);
	dsty = std::max(dsty, 16);
	HDC hDC = GetDC(NULL);
	m_hMenuBmp = CreateCompatibleBitmap(hDC, dstx, dsty);
	HDC hDCTemp = CreateCompatibleDC(hDC);
	ReleaseDC(NULL, hDC);
	HBITMAP hBitmapOld = (HBITMAP) ::SelectObject(hDCTemp, m_hMenuBmp);
	RECT rectBox = {0, 0, dstx, dsty};
	FillRect(hDCTemp, &rectBox, (HBRUSH)(COLOR_MENU + 1));
	DrawIconEx(hDCTemp, (dstx - 16) / 2, (dsty - 16) / 2, hIcon, 16, 16, 0, ::GetSysColorBrush(COLOR_MENU), DI_NORMAL);
	SelectObject(hDCTemp, hBitmapOld);
	DeleteDC(hDCTemp);
	DestroyIcon(hIcon);
}

DiskUsageTipExt::~DiskUsageTipExt(void)
{
	if (m_hMenuBmp)
	{
		DeleteObject(m_hMenuBmp);
		m_hMenuBmp = NULL;
	}

	InterlockedDecrement(&g_cDllRef);
}


#pragma region IUnknown

// Query to the interface the component supported.
IFACEMETHODIMP DiskUsageTipExt::QueryInterface(REFIID riid, void **ppv)
{
	static const QITAB qit[] =
	{
		QITABENT(DiskUsageTipExt, IContextMenu),
		QITABENT(DiskUsageTipExt, IShellExtInit),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppv);
}

// Increase the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) DiskUsageTipExt::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

// Decrease the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) DiskUsageTipExt::Release()
{
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (0 == cRef)
	{
		delete this;
	}

	return cRef;
}

#pragma endregion


#pragma region IShellExtInit

// determine whether dir is a mount point and get the volume name
// https://msdn.microsoft.com/en-us/library/windows/desktop/aa363940.aspx
static bool checkMountPoint(wchar_t *dir, wchar_t *volname, size_t volnamesize)
{
	// remove trailing '\\' if presents
	wchar_t *pend = dir + wcslen(dir);
	if (pend != dir && pend[-1] == L'\\')
		*(--pend) = 0;

	// check FILE_ATTRIBUTE_REPARSE_POINT
	DWORD attr = GetFileAttributesW(dir);
	if (attr == INVALID_FILE_ATTRIBUTES)	// error happened
		return false;
	if (!(attr & FILE_ATTRIBUTE_REPARSE_POINT))
		return false;
	//fprintf(stderr, "is reparse\n");

	// check IO_REPARSE_TAG_MOUNT_POINT
	WIN32_FIND_DATAW wfd;
	HANDLE hfind = FindFirstFileW(dir, &wfd);
	if (hfind == INVALID_HANDLE_VALUE)
		return false;
	bool ret = false;
	do
	{
		if (wfd.dwReserved0 == IO_REPARSE_TAG_MOUNT_POINT)
		{
			ret = true;
			break;
		}
	} while (FindNextFileW(hfind, &wfd));
	FindClose(hfind);
	if (!ret)
		return ret;

	// reparse the volume name
	*(pend++) = L'\\';
	*pend = 0;
	if (!GetVolumeNameForVolumeMountPointW(dir, volname, (DWORD)volnamesize))
		return false;
	*(--pend) = 0;

	return true;
}

static std::wstring formatsize(unsigned long long size)
{
	static const wchar_t * suffix[] = { L"", L"K", L"M", L"G", L"T", L"P", L"E", L"Z", L"Y" };
	size_t idx = 0;
	double dsize = (double)size;
	for (idx = 0; idx < sizeof(suffix) / sizeof(suffix[0]) - 1; ++idx)
	{
		if (dsize >= 1024)
			dsize /= 1024;
		else
			break;
	}
	wchar_t buf[40];
	_snwprintf_s(buf, 40, _TRUNCATE, L"%0.2f %sB", dsize, suffix[idx]);
	return buf;
}

// Initialize the context menu handler.
IFACEMETHODIMP DiskUsageTipExt::Initialize(
	LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID)
{
	m_diskUsageTip[0] = 0;
	//wcscpy(&m_diskUsageTip[0], m_pszMenuText);
	if (NULL == pDataObj)
	{
		return E_INVALIDARG;
	}

	HRESULT hr = E_FAIL;

	FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	STGMEDIUM stm;

	// The pDataObj pointer contains the objects being acted upon. In this 
	// example, we get an HDROP handle for enumerating the selected files and 
	// folders.
	if (!SUCCEEDED(pDataObj->GetData(&fe, &stm)))
		return hr;

	// Get an HDROP handle.
	HDROP hDrop = static_cast<HDROP>(GlobalLock(stm.hGlobal));
	if (!hDrop)
	{
		ReleaseStgMedium(&stm);
		return hr;
	}

	// Determine how many files are involved in this operation. This 
	// code sample displays the custom context menu item when only 
	// one file is selected.
	if (DragQueryFileW(hDrop, 0xFFFFFFFF, NULL, 0) == 1 &&
			0 != DragQueryFileW(hDrop, 0, m_szSelectedFile, ARRAYSIZE(m_szSelectedFile)))
	{
		//hr = S_OK;
		wchar_t volname[MAX_PATH];
		volname[0] = 0;
		DWORD spc, bps, fs, ts;
		if (!checkMountPoint(m_szSelectedFile, volname, MAX_PATH))
		{
			volname[0] = 0;
			size_t dlen = wcslen(m_szSelectedFile);
			if ((dlen == 2 || dlen == 3) && m_szSelectedFile[1] == L':' && (dlen == 2 || m_szSelectedFile[1] == L'\\'))
				wcsncpy(volname, m_szSelectedFile, MAX_PATH);
		}
		if (*volname && GetDiskFreeSpaceW(volname, &spc, &bps, &fs, &ts))
		{
			std::wstring tb = formatsize((unsigned long long)ts * spc * bps);
			std::wstring fb = formatsize((unsigned long long)fs * spc * bps);
			// L"%s free / %s total (%0.1f%%)"
			_snwprintf_s(&m_diskUsageTip[0], m_diskUsageTip.size(), _TRUNCATE, m_pszMenuText,
				fb.c_str(), tb.c_str(), (double)fs / ts * 100);
			hr = S_OK;
		}
	}

	GlobalUnlock(stm.hGlobal);
	ReleaseStgMedium(&stm);

	// If any value other than S_OK is returned from the method, the context 
	// menu item is not displayed.
	return hr;
}

#pragma endregion


#pragma region IContextMenu

//
//   FUNCTION: DiskUsageTipExt::QueryContextMenu
//
//   PURPOSE: The Shell calls IContextMenu::QueryContextMenu to allow the 
//            context menu handler to add its menu items to the menu. It 
//            passes in the HMENU handle in the hmenu parameter. The 
//            indexMenu parameter is set to the index to be used for the 
//            first menu item that is to be added.
//
IFACEMETHODIMP DiskUsageTipExt::QueryContextMenu(
	HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
	// If uFlags include CMF_DEFAULTONLY then we should not do anything.
	if (CMF_DEFAULTONLY & uFlags)
	{
		return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));
	}

	if (m_diskUsageTip[0] == 0)
		return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));

	// Add a separator.
	MENUITEMINFO sep = { sizeof(sep) };
	sep.fMask = MIIM_TYPE;
	sep.fType = MFT_SEPARATOR;
	if (!InsertMenuItemW(hMenu, indexMenu, TRUE, &sep))
	{
		return HRESULT_FROM_WIN32(GetLastError());
	}

	// Add the menu item.
	MENUITEMINFO mii = { sizeof(mii) };
	mii.fMask = MIIM_BITMAP | MIIM_STRING | MIIM_FTYPE | MIIM_ID | MIIM_STATE;
	mii.wID = idCmdFirst + IDM_DETAIL;
	mii.fType = MFT_STRING;
	mii.dwTypeData = &m_diskUsageTip[0];
	mii.fState = MFS_ENABLED;
	mii.hbmpItem = static_cast<HBITMAP>(m_hMenuBmp);
	if (!InsertMenuItemW(hMenu, indexMenu + 1, TRUE, &mii))
	{
		return HRESULT_FROM_WIN32(GetLastError());
	}

	// Add a separator.
	//MENUITEMINFO sep = { sizeof(sep) };
	//sep.fMask = MIIM_TYPE;
	//sep.fType = MFT_SEPARATOR;
	if (!InsertMenuItemW(hMenu, indexMenu + 2, TRUE, &sep))
	{
		return HRESULT_FROM_WIN32(GetLastError());
	}

	// Return an HRESULT value with the severity set to SEVERITY_SUCCESS. 
	// Set the code value to the offset of the largest command identifier 
	// that was assigned, plus one (1).
	return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(IDM_DETAIL + 1));
}


//
//   FUNCTION: DiskUsageTipExt::InvokeCommand
//
//   PURPOSE: This method is called when a user clicks a menu item to tell 
//            the handler to run the associated command. The lpcmi parameter 
//            points to a structure that contains the needed information.
//
IFACEMETHODIMP DiskUsageTipExt::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
	BOOL fUnicode = FALSE;

	// Determine which structure is being passed in, CMINVOKECOMMANDINFO or 
	// CMINVOKECOMMANDINFOEX based on the cbSize member of lpcmi. Although 
	// the lpcmi parameter is declared in Shlobj.h as a CMINVOKECOMMANDINFO 
	// structure, in practice it often points to a CMINVOKECOMMANDINFOEX 
	// structure. This struct is an extended version of CMINVOKECOMMANDINFO 
	// and has additional members that allow Unicode strings to be passed.
	if (pici->cbSize == sizeof(CMINVOKECOMMANDINFOEX))
	{
		if (pici->fMask & CMIC_MASK_UNICODE)
		{
			fUnicode = TRUE;
		}
	}

	// Determines whether the command is identified by its offset or verb.
	// There are two ways to identify commands:
	// 
	//   1) The command's verb string 
	//   2) The command's identifier offset
	// 
	// If the high-order word of lpcmi->lpVerb (for the ANSI case) or 
	// lpcmi->lpVerbW (for the Unicode case) is nonzero, lpVerb or lpVerbW 
	// holds a verb string. If the high-order word is zero, the command 
	// offset is in the low-order word of lpcmi->lpVerb.

	// For the ANSI case, if the high-order word is not zero, the command's 
	// verb string is in lpcmi->lpVerb. 
	if (HIWORD(pici->lpVerb))
	{
		return E_INVALIDARG;
		// Is the verb supported by this context menu extension?
		if (StrCmpIA(pici->lpVerb, m_pszVerb) == 0)
		{
			OnShowDetail(pici->hwnd);
		}
		else
		{
			// If the verb is not recognized by the context menu handler, it 
			// must return E_FAIL to allow it to be passed on to the other 
			// context menu handlers that might implement that verb.
			return E_FAIL;
		}
	}

	// For the Unicode case, if the high-order word is not zero, the 
	// command's verb string is in lpcmi->lpVerbW. 
	else if (fUnicode && HIWORD(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW))
	{
		// Is the verb supported by this context menu extension?
		if (StrCmpIW(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW, m_pwszVerb) == 0)
		{
			OnShowDetail(pici->hwnd);
		}
		else
		{
			// If the verb is not recognized by the context menu handler, it 
			// must return E_FAIL to allow it to be passed on to the other 
			// context menu handlers that might implement that verb.
			return E_FAIL;
		}
	}

	// If the command cannot be identified through the verb string, then 
	// check the identifier offset.
	else
	{
		// Is the command identifier offset supported by this context menu 
		// extension?
		if (LOWORD(pici->lpVerb) == IDM_DETAIL)
		{
			OnShowDetail(pici->hwnd);
		}
		else
		{
			// If the verb is not recognized by the context menu handler, it 
			// must return E_FAIL to allow it to be passed on to the other 
			// context menu handlers that might implement that verb.
			return E_FAIL;
		}
	}

	return S_OK;
}


//
//   FUNCTION: CFileContextMenuExt::GetCommandString
//
//   PURPOSE: If a user highlights one of the items added by a context menu 
//            handler, the handler's IContextMenu::GetCommandString method is 
//            called to request a Help text string that will be displayed on 
//            the Windows Explorer status bar. This method can also be called 
//            to request the verb string that is assigned to a command. 
//            Either ANSI or Unicode verb strings can be requested. This 
//            example only implements support for the Unicode values of 
//            uFlags, because only those have been used in Windows Explorer 
//            since Windows 2000.
//
IFACEMETHODIMP DiskUsageTipExt::GetCommandString(UINT_PTR idCommand,
	UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
	HRESULT hr = E_INVALIDARG;

	if (idCommand == IDM_DETAIL)
	{
		switch (uFlags)
		{
		case GCS_HELPTEXTW:
			// Only useful for pre-Vista versions of Windows that have a 
			// Status bar.
			hr = StringCchCopyW(reinterpret_cast<PWSTR>(pszName), cchMax,
				m_pwszVerbHelpText);
			break;

		case GCS_VERBW:
			// GCS_VERBW is an optional feature that enables a caller to 
			// discover the canonical name for the verb passed in through 
			// idCommand.
			hr = StringCchCopyW(reinterpret_cast<PWSTR>(pszName), cchMax,
				m_pwszVerbCanonicalName);
			break;

		default:
			hr = S_OK;
		}
	}

	// If the command (idCommand) is not supported by this context menu 
	// extension handler, return E_INVALIDARG.

	return hr;
}

#pragma endregion

// wprintf on a wchar_t vector
int vecwprintf(std::vector<wchar_t> &buf, size_t &pos, const wchar_t *format, ...)
{
	while (buf.size() <= pos + wcslen(format))
		buf.resize(buf.size() / 2 + buf.size() + 10);
	int res = -1;
	while (res < 0)
	{
		va_list vlist;
		va_start(vlist, format);
		res = _vsnwprintf_s(&buf[pos], buf.size() - pos, buf.size() - pos - 1, format, vlist);
		va_end(vlist);
		if (res < 0)
			buf.resize(buf.size() / 2 + buf.size());
	}
	pos += res;
	return 0;
}

void GetVolumePaths(wchar_t *volname, std::vector<std::wstring> &paths)
{
	paths.clear();
	DWORD charcnt = MAX_PATH + 1;
	wchar_t *names = new wchar_t[charcnt];
	BOOL success = FALSE;

	if (!(success = GetVolumePathNamesForVolumeNameW(volname, names, charcnt, &charcnt)) &&
			GetLastError() == ERROR_MORE_DATA)	// insufficient buffer
	{
		// Try again with the new suggested size.
		delete []names;
		names = new wchar_t[charcnt];
		success = GetVolumePathNamesForVolumeNameW(volname, names, charcnt, &charcnt);
	}

	if (success)
	{
		//  Extract the various paths.
		for (wchar_t *pname = names; pname[0] != '\0';)
		{
			const wchar_t *cname = pname;
			pname += wcslen(pname) + 1;
			if (pname[-2] == '\\')
				pname[-2] = 0;
			paths.push_back(cname);
		}
	}

	delete[] names;
	names = NULL;
	return;
}

void DiskUsageTipExt::OnShowDetail(HWND hWnd)
{
	bool verbose = false;
	std::vector<wchar_t> outbuf(300);
	size_t outpos = 0;

	//  Enumerate all volumes in the system.
	wchar_t  volname[MAX_PATH] = L"";
	HANDLE FindHandle = FindFirstVolumeW(volname, ARRAYSIZE(volname));
	if (FindHandle == INVALID_HANDLE_VALUE)
		return;
	do
	{
		//  Check the \\?\ prefix and remove the trailing backslash.
		size_t Index = wcslen(volname) - 1;
		if (wcsncmp(volname, L"\\\\?\\", 4) || volname[Index] != L'\\')
		{
			fwprintf(stderr, L"FindFirstVolume/FindNextVolume returned a bad path: %s\n", volname);
			continue;
		}

		//  QueryDosDevice does not allow a trailing backslash, so temporarily remove it.
		WCHAR DeviceName[MAX_PATH] = L"";
		volname[Index] = L'\0';
		DWORD CharCount = QueryDosDeviceW(&volname[4], DeviceName, ARRAYSIZE(DeviceName));
		volname[Index] = '\\';
		if (CharCount == 0)
		{
			fwprintf(stderr, L"QueryDosDevice failed with error code %d\n", (int)GetLastError());
			continue;
		}

		// Type
		static const wchar_t *typenames[] = { L"UNKNOWN", L"ERROR", L"RemovableMedia", L"FixedMedia", L"Remote", L"CDROM", L"RAM-disk" };
		UINT type = GetDriveTypeW(volname);
		if (type >= sizeof(typenames) / sizeof(typenames[0]))
			type = 1;
		if (type == DRIVE_REMOVABLE || type == DRIVE_CDROM || type == DRIVE_FIXED || type == DRIVE_REMOTE || type == DRIVE_RAMDISK)
		{
			// Show a '>' for current selected volume
			std::vector<std::wstring> paths;
			GetVolumePaths(volname, paths);
			const wchar_t *indicator = L"\x2001";
			for (auto i = paths.begin(); i != paths.end(); ++i)
			{
				if (m_szSelectedFile == *i)
				{
					indicator = L"->";
					break;
				}
			}
			vecwprintf(outbuf, outpos, L"%s ", indicator);

			// volume label
			wchar_t label[MAX_PATH + 1] = L"";
			wchar_t filesystem[MAX_PATH + 1] = L"";
			if (type == DRIVE_FIXED || type == DRIVE_REMOTE || type == DRIVE_RAMDISK)
			{
				if (!GetVolumeInformationW(volname, label, MAX_PATH + 1, NULL, NULL, NULL, filesystem, MAX_PATH + 1))
				{
					fwprintf(stderr, L"GetVolumeInformation failed.\n");
					label[0] = filesystem[0] = 0;
				}
				vecwprintf(outbuf, outpos, L"%s ", label);
			}
			else	// just show type for DRIVE_REMOVABLE and DRIVE_CDROM
				vecwprintf(outbuf, outpos, L"%s ", typenames[type]);

			// volume drive letters / mounted paths
			vecwprintf(outbuf, outpos, L"(");
			for (auto i = paths.begin(); i != paths.end(); ++i)
				vecwprintf(outbuf, outpos, L"%s%s", i == paths.begin() ? L"" : L"\x2000", i->c_str());
			vecwprintf(outbuf, outpos, L") ");

			// file system
			if (type == DRIVE_FIXED || type == DRIVE_REMOTE || type == DRIVE_RAMDISK)
				vecwprintf(outbuf, outpos, L"(%s) ", type == DRIVE_FIXED ? filesystem : typenames[type]);
			//vecwprintf(outbuf, outpos, L"\n");

			// free space
			if (type == DRIVE_FIXED)
			{
				// sizes
				DWORD spc, bps, fs, ts;
				if (GetDiskFreeSpaceW(volname, &spc, &bps, &fs, &ts))
				{
					std::wstring tb = formatsize((unsigned long long)ts * spc * bps);
					std::wstring ub = formatsize((unsigned long long)(ts - fs) * spc * bps);
					std::wstring fb = formatsize((unsigned long long)fs * spc * bps);
					vecwprintf(outbuf, outpos,
						L"\x2003%s\x3000%llu\n"
						L"\x2003\x2003\x2003"L"Free:\x2000\x3000%0.2f%%\x3000%s\x3000%llu\n"
						L"\x2003\x2003\x2003Used:\x3000%0.2f%%\x3000%s\x3000%llu\n",
						tb.c_str(), (unsigned long long)(ts - fs) * spc * bps,
						(float)fs / ts * 100, fb.c_str(), (unsigned long long)fs * spc * bps, 
						(float)(ts - fs) / ts * 100, ub.c_str(), (unsigned long long)ts * spc * bps);
				}
			}
		}

		if (verbose)
		{
			vecwprintf(outbuf, outpos, L"\x2003\x2003\x2003"L"Device Name: %s\n", DeviceName);
			vecwprintf(outbuf, outpos, L"\x2003\x2003\x2003"L"Volume name: %s\n", volname);
		}
		//vecwprintf(outbuf, outpos, L"\n");

		//  Move on to the next volume.
	} while(FindNextVolumeW(FindHandle, volname, ARRAYSIZE(volname)));

	FindVolumeClose(FindHandle);
	FindHandle = INVALID_HANDLE_VALUE;
	if (outbuf[outpos] == '\n')
		outbuf[outpos--] = 0;	// Remove last '\n'

	static const wchar_t detailCap[] = L"Disk Usage %s";
	size_t capbuflen = sizeof(detailCap) / sizeof(detailCap[0]) + wcslen(m_szSelectedFile);
	wchar_t *capbuf = new wchar_t[capbuflen];
	_snwprintf_s(capbuf, capbuflen, _TRUNCATE, detailCap, m_szSelectedFile);
	MessageBoxW(hWnd, &outbuf[0], capbuf, MB_OK | MB_ICONINFORMATION);
	delete[] capbuf;
}


