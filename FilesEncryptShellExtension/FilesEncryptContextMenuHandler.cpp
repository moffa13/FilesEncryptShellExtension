#include "FilesEncryptContextMenuHandler.h"
#include <sstream>

extern HINSTANCE g_hInstance;

FilesEncryptContextMenuHandler::FilesEncryptContextMenuHandler() : _objRefCount(1) {
	InterlockedIncrement(&g_dllCount);
}

HRESULT FilesEncryptContextMenuHandler::Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject *pdtobj, HKEY hkeyProgID) {

	HRESULT hr = E_INVALIDARG;
	if (NULL == pdtobj)
	{
		return hr;
	}

	FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	STGMEDIUM stm = {};

	// pDataObj contains the objects being acted upon. In this example, 
	// we get an HDROP handle for enumerating the selected files.
	if (pdtobj->GetData(&fe, &stm) == S_OK) {
		// Get an HDROP handle.
		HDROP hDrop = static_cast<HDROP>(GlobalLock(stm.hGlobal));
		if (hDrop != NULL) {
			// Determine how many files are involved in this operation.
			UINT nFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
			if (nFiles != 0) {
				_selectedFiles.clear();

				//Enumerates the selected files and directories.
				for (UINT i = 0; i < nFiles; i++) {
					// Get the next filename.
					int size = DragQueryFile(hDrop, i, NULL, 0) + 1;
					std::wstring str;
					str.resize(size);
					if (DragQueryFile(hDrop, i, &str[0], size) == 0)
						continue;

					_selectedFiles.push_back(str);
				}
				hr = S_OK;
			}

			GlobalUnlock(stm.hGlobal);
		}

		ReleaseStgMedium(&stm);
	}

	return hr;
}

DWORD FilesEncryptContextMenuHandler::filesState(){
	std::wstring exe = getFilesEncryptExecutable();
	
	std::wstring params = L"getState " + getFilesAsString();

	SHELLEXECUTEINFO infos = { 0 };
	infos.cbSize = sizeof(SHELLEXECUTEINFO);
	infos.fMask = SEE_MASK_NOCLOSEPROCESS;
	infos.hwnd = NULL;
	infos.lpVerb = NULL;
	infos.lpFile = exe.c_str();
	infos.lpParameters = params.c_str();
	infos.nShow = SW_HIDE;
	infos.hInstApp = NULL;

	ShellExecuteEx(&infos);
	WaitForSingleObject(infos.hProcess, INFINITE);

	DWORD code;
	GetExitCodeProcess(infos.hProcess, &code);
	return code;
}

std::wstring FilesEncryptContextMenuHandler::getFilesEncryptExecutable() {
	wchar_t filename[MAX_PATH] = { 0 };
	GetModuleFileName((HMODULE)g_hInstance, filename, MAX_PATH);
	std::wstring str = filename;
	std::wstring exe = str.substr(0, str.find_last_of('\\')) + L"\\FilesEncrypt.exe";
	return exe;
}

std::wstring FilesEncryptContextMenuHandler::getFilesAsString() {
	std::basic_stringstream<wchar_t> ss;

	for (std::vector<std::wstring>::iterator it = _selectedFiles.begin(); it != _selectedFiles.end(); ++it) {
		ss << "\"";
		ss.write(it->c_str(), it->size() - 1);
		ss << L"\" ";
	}

	std::wstring args = ss.str();
	args = args.substr(0, args.size() - 1);
	return args;
}

HRESULT FilesEncryptContextMenuHandler::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags) {

	if (uFlags & CMF_DEFAULTONLY)
		return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);

	DWORD filesState = this->filesState();

	UINT index = indexMenu;

	if (filesState == 1 || filesState == 3) {
		MENUITEMINFO encryptItem;
		encryptItem.cbSize = sizeof(MENUITEMINFO);
		encryptItem.fMask = MIIM_STRING | MIIM_ID;
		encryptItem.dwTypeData = L"Crypter avec FilesEncrypt";
		encryptItem.wID = idCmdFirst + CMD_ENCRYPT;
		_encryptId = encryptItem.wID;
		InsertMenuItem(hmenu, index++, TRUE, &encryptItem);
	}
	
	if (filesState == 2 || filesState == 3) {
		MENUITEMINFO decryptItem;
		decryptItem.cbSize = sizeof(MENUITEMINFO);
		decryptItem.fMask = MIIM_STRING | MIIM_ID;
		decryptItem.dwTypeData = L"Décrypter avec FilesEncrypt";
		decryptItem.wID = idCmdFirst + CMD_DECRYPT;
		_decryptId = decryptItem.wID;
		InsertMenuItem(hmenu, index++, TRUE, &decryptItem);
	}
		
	return MAKE_HRESULT(SEVERITY_SUCCESS, 0, CMD_LAST);
}

HRESULT FilesEncryptContextMenuHandler::InvokeCommand(CMINVOKECOMMANDINFO *pici) {
	CMINVOKECOMMANDINFOEX* cmd = (CMINVOKECOMMANDINFOEX*)pici;

	std::wstring exe = getFilesEncryptExecutable();

	WORD verb;

	if (!HIWORD(cmd->lpVerbW)) { // If verb is int
		verb = LOWORD(pici->lpVerb);
		if (verb != CMD_DECRYPT && verb != CMD_ENCRYPT) return E_FAIL;
	} else {
		return E_FAIL;
	}

	std::wstring args = std::wstring{ (verb == CMD_DECRYPT ? L"decrypt" : L"encrypt") } +L" " + getFilesAsString();

	ShellExecute(NULL, L"open", exe.c_str(), args.c_str(), NULL, SW_SHOWDEFAULT);
	return S_OK;
}

HRESULT FilesEncryptContextMenuHandler::GetCommandString(UINT_PTR idCmd, UINT uType, UINT *pReserved, CHAR *pszName, UINT cchMax) {
	return S_OK;
}

HRESULT FilesEncryptContextMenuHandler::QueryInterface(REFIID riid, void **ppvObject) {
	if (!ppvObject) return E_POINTER;
	*ppvObject = NULL;
	if (IsEqualIID(riid, IID_IUnknown)) {
		*ppvObject = this;
		AddRef();
		return S_OK;
	} else if (IsEqualIID(riid, IID_IShellExtInit)) {
		*ppvObject = (IShellExtInit*)this;
		AddRef();
		return S_OK;
	} else if (IsEqualIID(riid, IID_IContextMenu)) {
		*ppvObject = (IContextMenu*)this;
		AddRef();
		return S_OK;
	} else {
		return E_NOINTERFACE;
	}
}

ULONG FilesEncryptContextMenuHandler::AddRef(){
	return InterlockedIncrement(&_objRefCount);
}

ULONG FilesEncryptContextMenuHandler::Release() {
	auto dec = InterlockedDecrement(&_objRefCount);
	if (dec < 1) {
		delete this;
	}
	return dec;
}

FilesEncryptContextMenuHandler::~FilesEncryptContextMenuHandler() {
	InterlockedDecrement(&g_dllCount);
}
