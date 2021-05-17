#include <windows.h>
#include <tchar.h>
#include <initguid.h>
#include "../IExample/IExample.h"


static const TCHAR	OurDllName[] = _T("IExample.dll");
static const TCHAR	ObjectDescription[] = _T("IExample COM component");
static const TCHAR	FileDlgTitle[] = _T("Locate IExample.dll to register it");

static const TCHAR	FileDlgExt[] = _T("DLL files\000*.dll\000\000");
static const TCHAR	ClassKeyName[] = _T("Software\\Classes");
static const TCHAR	CLSID_Str[] = _T("CLSID");
static const TCHAR	InprocServer32Name[] = _T("InprocServer32");
static const TCHAR	ThreadingModel[] = _T("ThreadingModel");
static const TCHAR	BothStr[] = _T("both");
static const TCHAR	GUID_Format[] = _T("{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}");

static void stringFromCLSID(LPTSTR buffer, REFCLSID ri)
{
	wsprintf(buffer, &GUID_Format[0],
		((REFCLSID)ri)->Data1, ((REFCLSID)ri)->Data2, ((REFCLSID)ri)->Data3, ((REFCLSID)ri)->Data4[0],
		((REFCLSID)ri)->Data4[1], ((REFCLSID)ri)->Data4[2], ((REFCLSID)ri)->Data4[3],
		((REFCLSID)ri)->Data4[4], ((REFCLSID)ri)->Data4[5], ((REFCLSID)ri)->Data4[6],
		((REFCLSID)ri)->Data4[7]);
}

static void cleanup(void)
{
	HKEY		rootKey;
	HKEY		hKey;
	HKEY		hKey2;
	TCHAR		buffer[39];

	stringFromCLSID(&buffer[0], (REFCLSID)(&CLSID_IExample));

	// Open "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Classes"
	if (!RegOpenKeyEx(HKEY_LOCAL_MACHINE, &ClassKeyName[0], 0, KEY_WRITE, &rootKey))
	{
		// Delete our CLSID key and everything under it
		if (!RegOpenKeyEx(rootKey, &CLSID_Str[0], 0, KEY_ALL_ACCESS, &hKey))
		{
			if (!RegOpenKeyEx(hKey, &buffer[0], 0, KEY_ALL_ACCESS, &hKey2))
			{
				RegDeleteKey(hKey2, &InprocServer32Name[0]);

				RegCloseKey(hKey2);
				RegDeleteKey(hKey, &buffer[0]);
			}

			RegCloseKey(hKey);
		}

		RegCloseKey(rootKey);
	}
}

int WINAPI WinMain(HINSTANCE hinstExe, HINSTANCE hinstPrev, LPSTR lpszCmdLine, int nCmdShow)
{
	int				result;
	TCHAR			filename[MAX_PATH];

	{
	OPENFILENAME	ofn;

	// Pick out where our DLL is located. We need to know its location in
	// order to register it as a COM component
	lstrcpy(&filename[0], &OurDllName[0]);
	ZeroMemory(&ofn, sizeof(OPENFILENAME));
	ofn.lStructSize = sizeof(OPENFILENAME);
	ofn.lpstrFilter = &FileDlgExt[0];
	ofn.lpstrFile = &filename[0];
	ofn.nMaxFile = MAX_PATH;
	ofn.lpstrTitle = &FileDlgTitle[0];
	ofn.Flags = OFN_FILEMUSTEXIST|OFN_EXPLORER|OFN_PATHMUSTEXIST;
	result = GetOpenFileName(&ofn);
	}

	if (result > 0)
	{
		HKEY		rootKey;
		HKEY		hKey;
		HKEY		hKey2;
		HKEY		hkExtra;
		TCHAR		buffer[39];
		DWORD		disposition;

		// Assume an error
		result = 1;

		// Open "HKEY_LOCAL_MACHINE\Software\Classes"
		if (!RegOpenKeyEx(HKEY_LOCAL_MACHINE, &ClassKeyName[0], 0, KEY_WRITE, &rootKey))
		{
			// Open "HKEY_LOCAL_MACHINE\Software\Classes\CLSID"
			if (!RegOpenKeyEx(rootKey, &CLSID_Str[0], 0, KEY_ALL_ACCESS, &hKey))
			{
				stringFromCLSID(&buffer[0], (REFCLSID)(&CLSID_IExample));
				if (!RegCreateKeyEx(hKey, &buffer[0], 0, 0, REG_OPTION_NON_VOLATILE, KEY_WRITE, 0, &hKey2, &disposition))
				{
					RegSetValueEx(hKey2, 0, 0, REG_SZ, (const BYTE *)&ObjectDescription[0], sizeof(ObjectDescription));
					if (!RegCreateKeyEx(hKey2, &InprocServer32Name[0], 0, 0, REG_OPTION_NON_VOLATILE, KEY_WRITE, 0, &hkExtra, &disposition))
					{
						if (!RegSetValueEx(hkExtra, 0, 0, REG_SZ, (const BYTE *)&filename[0], lstrlen(&filename[0]) + 1))
						{
							if (!RegSetValueEx(hkExtra, &ThreadingModel[0], 0, REG_SZ, (const BYTE *)&BothStr[0], sizeof(BothStr)))
							{
								result = 0;
								MessageBox(0, "Successfully registered IExample.DLL as a COM component.", &ObjectDescription[0], MB_OK);
							}
						}
						RegCloseKey(hkExtra);
					}
					RegCloseKey(hKey2);
				}
				RegCloseKey(hKey);
			}
			RegCloseKey(rootKey);
		}

		if (result)
		{
			cleanup();
			MessageBox(0, "Failed to register IExample.DLL as a COM component, try runing as administrator.", &ObjectDescription[0], MB_OK|MB_ICONEXCLAMATION);
		}
	}

	return(0);
}
