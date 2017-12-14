#pragma once
#include <Windows.h>
#include <vector>
#include <Shobjidl.h>

extern UINT g_dllCount;

class FilesEncryptContextMenuHandler : public IShellExtInit, IContextMenu, IUnknown {
private:
	std::vector<std::wstring> m_selectedFiles;
	UINT_PTR m_idCmd;
protected:
	DWORD _objRefCount;
	~FilesEncryptContextMenuHandler();
public:

	FilesEncryptContextMenuHandler();

	// IShellExtInit
	HRESULT __stdcall Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject *pdtobj, HKEY hkeyProgID) override;

	// IContextMenu
	HRESULT __stdcall QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags) override;
	HRESULT __stdcall InvokeCommand(CMINVOKECOMMANDINFO *pici) override;
	HRESULT __stdcall GetCommandString(UINT_PTR idCmd, UINT uType, UINT *pReserved, CHAR *pszName, UINT cchMax) override;

	// IUnknown
	HRESULT  __stdcall QueryInterface(REFIID riid, void **ppvObject) override;
	ULONG __stdcall AddRef() override;
	ULONG __stdcall Release() override;
};

