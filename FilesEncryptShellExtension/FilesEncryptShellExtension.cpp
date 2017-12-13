#include <Windows.h>

BOOL WINAPI DllMain(
	HINSTANCE hinstDLL,
	DWORD     fdwReason,
	LPVOID    lpvReserved
) {

	switch (fdwReason) {
		case DLL_PROCESS_ATTACH:
			break;
	}

	return TRUE;
}

HRESULT WINAPI DllRegisterServer() {
	return E_NOTIMPL;
}

HRESULT WINAPI DllUnregisterServer() {
	return E_NOTIMPL;
}

HRESULT WINAPI DllCanUnloadNow() {
	return E_NOTIMPL;
}

HRESULT WINAPI DllGetClassObject(
	REFCLSID rclsid,
	REFIID   riid,
	LPVOID   *ppv
) {
	return E_NOTIMPL;
}