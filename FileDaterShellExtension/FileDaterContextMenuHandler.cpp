#include "stdafx.h"
#include <strsafe.h>
#include <shellapi.h>
#include <stdexcept>
#include "FileDaterShellExtension.h"
#include "FileDaterContextMenuHandler.h"

FileDaterContextMenuHandler::FileDaterContextMenuHandler() : m_dwRefCount(1)
{
	InterlockedIncrement(&g_cObjCount);
}

ULONG FileDaterContextMenuHandler::AddRef()
{
	return InterlockedIncrement(&m_dwRefCount);
}

ULONG FileDaterContextMenuHandler::Release() {
	ULONG iValue;
	iValue = InterlockedDecrement(&m_dwRefCount);
	if (iValue < 1)
		delete this;
	return iValue;
}

HRESULT FileDaterContextMenuHandler::QueryInterface(REFIID riid, void **ppvObject) {
	if (!ppvObject)
		return E_POINTER;

	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown)) {
		*ppvObject = this;
		this->AddRef();
		return S_OK;
	}
	else if (IsEqualIID(riid, IID_IContextMenu)) {
		*ppvObject = (IContextMenu*)this;
		this->AddRef();
		return S_OK;
	}
	else if (IsEqualIID(riid, IID_IShellExtInit)) {
		*ppvObject = (IShellExtInit*)this;
		this->AddRef();
		return S_OK;
	}
	else
		return E_NOINTERFACE;
}

HRESULT FileDaterContextMenuHandler::Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject *pdtobj, HKEY hkeyProgID)
{
	HRESULT hResult = E_FAIL;
	FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
	STGMEDIUM stm;
	wchar_t wszSelectedFile[MAX_PATH * sizeof(wchar_t)];

	if (NULL == pdtobj)
		return E_INVALIDARG;

	if (SUCCEEDED(pdtobj->GetData(&fe, &stm)))
	{
		HDROP hDrop = static_cast<HDROP>(GlobalLock(stm.hGlobal));
		if (hDrop != NULL)
		{
			UINT nFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
			if (nFiles == 1)
			{
				if (0 != DragQueryFile(hDrop, 0, wszSelectedFile, ARRAYSIZE(wszSelectedFile))) {
					try {
						m_objFileDater = new FileDater(wszSelectedFile);
						hResult = S_OK;
					}
					catch (...) {
					}
					
				}
			}
			GlobalUnlock(stm.hGlobal);
		}
		ReleaseStgMedium(&stm);
	}

	return hResult;
}

HRESULT FileDaterContextMenuHandler::GetCommandString(UINT_PTR idCmd, UINT uType, UINT *pReserved, CHAR *pszName, UINT cchMax)
{
	return E_NOTIMPL;
}

HRESULT FileDaterContextMenuHandler::InvokeCommand(CMINVOKECOMMANDINFO *pici)
{
	HRESULT hResult = E_INVALIDARG;
	if (!HIWORD(pici->lpVerb)) {
		UINT idCmd = LOWORD(pici->lpVerb);

		switch (idCmd) {
		case 0: // Rename
			return m_objFileDater->Rename();
		case 1: // Copy
			return m_objFileDater->Copy();
		}
	}

	return hResult;
}

HRESULT FileDaterContextMenuHandler::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
	MENUITEMINFO miiRename = {};
	MENUITEMINFO miiCopy = {};
	wchar_t wszRenameLabel[_MAX_PATH + 64], wszCopyLabel[_MAX_PATH + 64];

	if (uFlags & CMF_DEFAULTONLY)
		return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);

	StringCbPrintf(wszRenameLabel, sizeof(wszRenameLabel), L"Rename to \"%s\"", m_objFileDater->m_wszDstName);
	miiRename.cbSize = sizeof(MENUITEMINFO);
	miiRename.fMask = MIIM_STRING | MIIM_ID;
	miiRename.wID = idCmdFirst;
	miiRename.dwTypeData = wszRenameLabel;

	if (!InsertMenuItem(hmenu, 0, TRUE, &miiRename))
		return HRESULT_FROM_WIN32(GetLastError());

	StringCbPrintf(wszCopyLabel, sizeof(wszCopyLabel), L"Copy to \"%s\"", m_objFileDater->m_wszDstName);
	miiCopy.cbSize = sizeof(MENUITEMINFO);
	miiCopy.fMask = MIIM_STRING | MIIM_ID;
	miiCopy.wID = idCmdFirst + 1;
	miiCopy.dwTypeData = wszCopyLabel;

	if (!InsertMenuItem(hmenu, 1, TRUE, &miiCopy))
		return HRESULT_FROM_WIN32(GetLastError());

	return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 2);
}

FileDaterContextMenuHandler::~FileDaterContextMenuHandler()
{
	if (m_objFileDater)
		delete m_objFileDater;
	InterlockedDecrement(&g_cObjCount);
}
