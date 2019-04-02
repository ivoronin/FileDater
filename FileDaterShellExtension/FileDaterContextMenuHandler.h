#pragma once
#include "stdafx.h"
#include <Shobjidl.h>
#include "FileDater.h"

class FileDaterContextMenuHandler : public IShellExtInit, IContextMenu
{
private:
	FileDater *m_objFileDater;
protected:
	DWORD m_dwRefCount;
	~FileDaterContextMenuHandler();
public:
	FileDaterContextMenuHandler();

	// IUnknown
	ULONG STDMETHODCALLTYPE AddRef();
	ULONG STDMETHODCALLTYPE Release();
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppvObject);

	// IShellExtInit
	HRESULT STDMETHODCALLTYPE Initialize(PCIDLIST_ABSOLUTE pidlFolder, IDataObject *pdtobj, HKEY hkeyProgID);

	// IContextMenu
	HRESULT STDMETHODCALLTYPE GetCommandString(
		UINT_PTR idCmd,
		UINT     uType,
		UINT     *pReserved,
		CHAR     *pszName,
		UINT     cchMax
	);

	HRESULT STDMETHODCALLTYPE InvokeCommand(
		CMINVOKECOMMANDINFO *pici
	);

	HRESULT STDMETHODCALLTYPE QueryContextMenu(
		HMENU hmenu,
		UINT  indexMenu,
		UINT  idCmdFirst,
		UINT  idCmdLast,
		UINT  uFlags
	);
};

