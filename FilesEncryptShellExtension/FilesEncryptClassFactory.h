#pragma once
#include <Windows.h>

extern UINT g_dllCount;

class FilesEncryptClassFactory : public IClassFactory, IUnknown {

protected:
	DWORD _objRefCount;
	~FilesEncryptClassFactory();
public:
	FilesEncryptClassFactory();

	// IUnknown
	HRESULT  __stdcall QueryInterface(REFIID riid, void **ppvObject) override;
	ULONG __stdcall AddRef() override;
	ULONG __stdcall Release() override;

	// IClassFactory
	HRESULT __stdcall CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject) override;
	HRESULT __stdcall LockServer(BOOL fLock) override;
};