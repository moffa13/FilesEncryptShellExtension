#pragma once
#include <Windows.h>
#include <vector>
#include <Shobjidl.h>

extern UINT g_dllCount;

enum CMDS {
	CMD_FIRST = 0,
	CMD_ENCRYPT = CMD_FIRST,
	CMD_DECRYPT,
	CMD_LAST
};

class FilesEncryptContextMenuHandler : public IShellExtInit, IContextMenu, IUnknown {
private:
	std::vector<std::wstring> _selectedFiles;
	UINT _decryptId = 0;
	UINT _encryptId = 0;
	DWORD filesState();
	std::wstring getFilesEncryptExecutable();
	std::wstring getFilesAsString();
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

