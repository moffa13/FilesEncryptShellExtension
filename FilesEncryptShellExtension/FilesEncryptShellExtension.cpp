#include <Windows.h>
#include <ShlObj.h>
#include <string>
#include <sstream>
#include "FilesEncryptClassFactory.h"

#include "GUID.h"

#define ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_NO_CLOSE_REG(x) if(x != ERROR_SUCCESS) {return E_UNEXPECTED;}
#define ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_CLOSE_REG(x) if(x != ERROR_SUCCESS) {RegCloseKey(hKey); return E_UNEXPECTED;}

static HANDLE g_hInstance;
UINT g_dllCount = 0;
static const std::wstring dllName = L"FilesEncrypt";

BOOL WINAPI DllMain(
	HINSTANCE hinstDLL,
	DWORD     fdwReason,
	LPVOID    lpvReserved
) {

	switch (fdwReason) {
		case DLL_PROCESS_ATTACH:
			g_hInstance = hinstDLL;
			break;
	}

	return TRUE;
}

std::wstring GetDllPath() {
	wchar_t filename[MAX_PATH];
	GetModuleFileName((HMODULE)g_hInstance, filename, MAX_PATH);
	return std::wstring(filename);
}

std::wstring GetCLSID() {
	// Create the clsid path string
	LPOLESTR clsid;
	StringFromCLSID(FilesEncryptShellExtensionGUID, &clsid);
	CoTaskMemFree(clsid);
	return std::wstring(clsid);
}

std::wstring GetCLSIDPath() {
	return L"SOFTWARE\\Classes\\CLSID\\" + GetCLSID();
}

DWORD GetStringSizeBytes(const std::wstring& str) {
	return (str.size() + 1) * 2;
}

HRESULT WINAPI DllRegisterServer() {

	std::wstring subKey = GetCLSIDPath();

	HKEY hKey;
	LONG res = RegCreateKeyEx(HKEY_CURRENT_USER, subKey.c_str(), NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
	ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_CLOSE_REG(res);

	res = RegCreateKeyEx(hKey, L"InProcServer32", NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
	ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_CLOSE_REG(res);

	std::wstring path = GetDllPath();
	res = RegSetKeyValue(hKey, NULL, NULL, REG_SZ, path.c_str(), GetStringSizeBytes(path));
	ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_CLOSE_REG(res);

	std::wstring model = L"Apartment";
	res = RegSetKeyValue(hKey, NULL, L"ThreadingModel", REG_SZ, model.c_str(), GetStringSizeBytes(model));
	ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_CLOSE_REG(res);

	RegCloseKey(hKey);

	res = RegCreateKeyEx(HKEY_CURRENT_USER, (L"SOFTWARE\\Classes\\*\\shellex\\ContextMenuHandlers\\" + dllName).c_str(), NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
	ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_CLOSE_REG(res);

	std::wstring clsid = GetCLSID();
	res = RegSetKeyValue(hKey, NULL, NULL, REG_SZ, clsid.c_str(), GetStringSizeBytes(clsid));
	ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_CLOSE_REG(res);

	RegCloseKey(hKey);

	res = RegCreateKeyEx(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved", NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
	ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_CLOSE_REG(res);
	res = RegSetKeyValue(hKey, NULL, clsid.c_str(), REG_SZ, dllName.c_str(), GetStringSizeBytes(dllName));
	ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_CLOSE_REG(res);

	RegCloseKey(hKey);

	SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

	return S_OK;
}

HRESULT WINAPI DllUnregisterServer() {
	std::wstring clsidPath = GetCLSIDPath();
	LONG res = RegDeleteTree(HKEY_CURRENT_USER, clsidPath.c_str());
	//ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_NO_CLOSE_REG(res);
	res = RegDeleteKey(HKEY_CURRENT_USER, (L"SOFTWARE\\Classes\\*\\ShellEx\\ContextMenuHandlers\\" + dllName).c_str());
	//ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_NO_CLOSE_REG(res);
	HKEY hKey;
	RegOpenCurrentUser(KEY_ALL_ACCESS, &hKey);
	RegOpenKeyEx(hKey, L"Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved", 0, KEY_ALL_ACCESS, &hKey);
	res = RegDeleteValue(hKey, GetCLSID().c_str());
	RegCloseKey(hKey);
	//ERROR_SUCCESS_OR_RETURN_E_UNEXPECTED_NO_CLOSE_REG(res);
	SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
	return S_OK;
}

HRESULT WINAPI DllCanUnloadNow() {
	return g_dllCount > 0 ? S_FALSE : S_OK;
}

HRESULT WINAPI DllGetClassObject(
	REFCLSID rclsid,
	REFIID   riid,
	LPVOID   *ppv
) {
	if (!ppv) return E_INVALIDARG;
	*ppv = NULL;

	if (!IsEqualCLSID(rclsid, FilesEncryptShellExtensionGUID)) return CLASS_E_CLASSNOTAVAILABLE;

	HRESULT hr = E_UNEXPECTED;

	FilesEncryptClassFactory *pFactory = new FilesEncryptClassFactory;
	if (pFactory) {
		hr = pFactory->QueryInterface(riid, ppv);
		pFactory->Release();
	}

	return S_OK;
}