#include "FilesEncryptContextMenuHandler.h"

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
				m_selectedFiles.clear();

				//Enumerates the selected files and directories.
				for (UINT i = 0; i < nFiles; i++) {
					// Get the next filename.
					int size = DragQueryFile(hDrop, i, NULL, 0) + 1;
					std::wstring str;
					str.resize(size);
					if (DragQueryFile(hDrop, i, &str[0], size) == 0)
						continue;

					m_selectedFiles.push_back(str);
				}
				hr = S_OK;
			}

			GlobalUnlock(stm.hGlobal);
		}

		ReleaseStgMedium(&stm);
	}

	return hr;
}

HRESULT FilesEncryptContextMenuHandler::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags) {

	if (uFlags & CMF_DEFAULTONLY)
		return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);


	MENUITEMINFO encryptItem;
	encryptItem.cbSize = sizeof(MENUITEMINFO);
	encryptItem.fMask = MIIM_STRING | MIIM_ID;
	encryptItem.dwTypeData = L"Crypter avec FilesEncrypt";
	encryptItem.wID = idCmdFirst;

	MENUITEMINFO decryptItem;
	encryptItem.cbSize = sizeof(MENUITEMINFO);
	encryptItem.fMask = MIIM_STRING | MIIM_ID;
	encryptItem.dwTypeData = L"Décrypter avec FilesEncrypt";
	encryptItem.wID = idCmdFirst + 1;

	InsertMenuItem(hmenu, 0, TRUE, &encryptItem);
	InsertMenuItem(hmenu, 1, TRUE, &decryptItem);

	return MAKE_HRESULT(SEVERITY_SUCCESS, 0, decryptItem.wID - idCmdFirst + 1);
}

HRESULT FilesEncryptContextMenuHandler::InvokeCommand(CMINVOKECOMMANDINFO *pici) {
	MessageBox(NULL, m_selectedFiles[0].c_str(), m_selectedFiles[1].c_str(), MB_OK);
	return S_OK;
}

HRESULT FilesEncryptContextMenuHandler::GetCommandString(UINT_PTR idCmd, UINT uType, UINT *pReserved, CHAR *pszName, UINT cchMax) {
	return E_NOTIMPL;
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
