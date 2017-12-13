#include "FilesEncryptClassFactory.h"
#include "FilesEncryptContextMenuHandler.h"
#include <Shobjidl.h>

FilesEncryptClassFactory::~FilesEncryptClassFactory() {
	InterlockedDecrement(&g_dllCount);
}

FilesEncryptClassFactory::FilesEncryptClassFactory() : _objRefCount(1) {
	InterlockedIncrement(&g_dllCount);
}

HRESULT FilesEncryptClassFactory::QueryInterface(REFIID riid, void **ppvObject) {
	if (!ppvObject) return E_POINTER;
	*ppvObject = NULL;
	if (IsEqualIID(riid, IID_IUnknown)) {
		*ppvObject = this;
		AddRef();
		return S_OK;
	} else if (IsEqualIID(riid, IID_IClassFactory)) {
		*ppvObject = (IClassFactory*)this;
		AddRef();
		return S_OK;
	} else {
		return E_NOINTERFACE;
	}
}

ULONG FilesEncryptClassFactory::AddRef() {
	return InterlockedIncrement(&_objRefCount);
}

ULONG FilesEncryptClassFactory::Release() {
	auto dec = InterlockedDecrement(&_objRefCount);
	if (dec < 1) {
		delete this;
	}
	return dec;
}

HRESULT FilesEncryptClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject) {
	if (!ppvObject) return E_INVALIDARG;
	if (pUnkOuter != NULL) return CLASS_E_NOAGGREGATION;

	HRESULT hr = E_UNEXPECTED;
	if (IsEqualIID(riid, IID_IShellExtInit) || IsEqualIID(riid, IID_IContextMenu)) {
		FilesEncryptContextMenuHandler *pHandler = new FilesEncryptContextMenuHandler;
		if (!pHandler) return E_OUTOFMEMORY;
		hr = pHandler->QueryInterface(riid, ppvObject);
		pHandler->Release();
	} else {
		hr = E_NOINTERFACE;
	}
	return hr;
}

HRESULT FilesEncryptClassFactory::LockServer(BOOL fLock) {
	return S_OK;
}
