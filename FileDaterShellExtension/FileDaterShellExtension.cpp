#include "stdafx.h"
#include <Objbase.h>
#include <Shlobj.h>
#include <string>
#include "FileDaterClassFactory.h"
#include "FileDaterShellExtension.h"

UINT g_cObjCount;

std::wstring GetShellExtensionGUIDString() {
	HRESULT hResult;
	wchar_t * wszGUIDString;

	/* Convert GUID to string */
	hResult = StringFromCLSID(CLSID_FileDaterShellExtension, &wszGUIDString);
	if (hResult != S_OK)
		return NULL;
	return std::wstring(wszGUIDString);
}

/* Return size of a wide-character string in bytes */
DWORD wcslenb(const wchar_t * wszString) {
	return (DWORD)(wcslen(wszString) + 1) * sizeof(wchar_t);
}

HRESULT __stdcall DllCanUnloadNow() {
	if (g_cObjCount > 0)
		return S_FALSE;
	else
		return S_OK;
}

HRESULT __stdcall DllGetClassObject(REFCLSID rclsid,
	REFIID   riid,
	LPVOID   *ppv) {
	FileDaterClassFactory *pFactory;
	HRESULT hResult;

	if (!IsEqualCLSID(rclsid, CLSID_FileDaterShellExtension))
		return CLASS_E_CLASSNOTAVAILABLE;

	if (!ppv)
		return E_INVALIDARG;

	hResult = E_UNEXPECTED;
	pFactory = new FileDaterClassFactory();
	if (pFactory != NULL) {
		hResult = pFactory->QueryInterface(riid, ppv);
		pFactory->Release();
	}

	return hResult;
}

HRESULT __stdcall DllRegisterServer() {
	HKEY hkGUID, hkInprocServer32, hkHandler;
	LSTATUS lStatus;

	wchar_t dllPath[MAX_PATH];
	std::wstring lpGUIDString, lpGUIDKey, lpHandlerKey;
	wchar_t wszApartment[] = L"Apartment";

	/*
	 * Prepare data
	 */

	 /* Get dll path */
	if (GetModuleFileName(g_hModule, dllPath, MAX_PATH) == 0)
		return E_UNEXPECTED;

	lpGUIDString = GetShellExtensionGUIDString();

	/*
	 * Create required registy keys
	 */
	lpGUIDKey = L"Software\\Classes\\CLSID\\" + lpGUIDString; // FIXME: free lpSubKey

	/* Create HKEY_LOCAL_MACHINE\Software\Classes\CLSID\{GUID}\ key */
	lStatus = RegCreateKeyEx(HKEY_LOCAL_MACHINE, lpGUIDKey.c_str(), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hkGUID, NULL);
	if (lStatus != ERROR_SUCCESS)
		return E_UNEXPECTED;

	/* Create InProcServer32 subkey */
	lStatus = RegCreateKeyEx(hkGUID, L"InProcServer32", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hkInprocServer32, NULL);
	if (lStatus != ERROR_SUCCESS)
		return E_UNEXPECTED;

	/* Close GUID key */
	lStatus = RegCloseKey(hkGUID);
	if (lStatus != ERROR_SUCCESS)
		return E_UNEXPECTED;

	/* Set value for InProcServer32 default subkey */
	lStatus = RegSetValueEx(hkInprocServer32, NULL, 0, REG_SZ, (BYTE*)&dllPath, wcslenb(dllPath));
	if (lStatus != ERROR_SUCCESS)
		return E_UNEXPECTED;

	/* Set value for InProcServer32 ThreadingModel subkey */
	lStatus = RegSetValueEx(hkInprocServer32, L"ThreadingModel", 0, REG_SZ, (BYTE*)&wszApartment, wcslenb(wszApartment));
	if (lStatus != ERROR_SUCCESS)
		return E_UNEXPECTED;

	/* Close InProcServer32 subkey */
	lStatus = RegCloseKey(hkInprocServer32);
	if (lStatus != ERROR_SUCCESS)
		return E_UNEXPECTED;

	/* Create handler key */
	lpHandlerKey = L"Software\\Classes\\*\\shellex\\ContextMenuHandlers\\FileDater";
	lStatus = RegCreateKeyEx(HKEY_LOCAL_MACHINE, lpHandlerKey.c_str(), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hkHandler, NULL);
	if (lStatus != ERROR_SUCCESS)
		return E_UNEXPECTED;

	/* Set value for InProcServer32 default subkey */
	const wchar_t * wszGUIDString = lpGUIDString.c_str();
	lStatus = RegSetValueEx(hkHandler, NULL, 0, REG_SZ, (BYTE*)wszGUIDString, wcslenb(wszGUIDString));
	if (lStatus != ERROR_SUCCESS)
		return E_UNEXPECTED;

	/* Close handler key */
	lStatus = RegCloseKey(hkHandler);
	if (lStatus != ERROR_SUCCESS)
		return E_UNEXPECTED;

	SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

	return S_OK;
}

HRESULT __stdcall DllUnregisterServer() {
	LSTATUS lStatus;
	std::wstring lpGUIDKey, lpInprocServer32Key, lpHandlerKey;

	lpGUIDKey = L"Software\\Classes\\CLSID\\" + GetShellExtensionGUIDString();
	lpInprocServer32Key = lpGUIDKey + L"\\InprocServer32";
	lpHandlerKey = L"Software\\Classes\\*\\shellex\\ContextMenuHandlers\\"  L"FileDater";

	lStatus = RegDeleteTree(HKEY_LOCAL_MACHINE, lpGUIDKey.c_str());
	if (lStatus != ERROR_SUCCESS)
		return E_UNEXPECTED;

	lStatus = RegDeleteTree(HKEY_LOCAL_MACHINE, lpHandlerKey.c_str());
	if (lStatus != ERROR_SUCCESS)
		return E_UNEXPECTED;

	SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
	return S_OK;
}