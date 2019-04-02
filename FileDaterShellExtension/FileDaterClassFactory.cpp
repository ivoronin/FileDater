#include "stdafx.h"
#include <guiddef.h>
#include "FileDaterClassFactory.h"
#include "FileDaterShellExtension.h"
#include "FileDaterContextMenuHandler.h"

FileDaterClassFactory::FileDaterClassFactory() : m_dwRefCount(1)
{
	InterlockedIncrement(&g_cObjCount);
}

ULONG FileDaterClassFactory::AddRef() {
	return InterlockedIncrement(&m_dwRefCount);
}

ULONG FileDaterClassFactory::Release() {
	ULONG iValue;
	iValue = InterlockedDecrement(&m_dwRefCount);
	if (iValue < 1)
	{
		delete this;
	}
	return iValue;
}

HRESULT FileDaterClassFactory::QueryInterface(REFIID riid, void **ppvObject) {
	if (!ppvObject)
		return E_POINTER;

	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown)) {
		*ppvObject = this;
		this->AddRef();
		return S_OK;
	}
	else if (IsEqualIID(riid, IID_IClassFactory)) {
		*ppvObject = (IClassFactory*)this;
		this->AddRef();
		return S_OK;
	}
	else {
		return E_NOINTERFACE;
	}
}

HRESULT FileDaterClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
{
	FileDaterContextMenuHandler *pHandler;
	HRESULT hResult;

	if (!ppvObject)
		return E_INVALIDARG;

	if (pUnkOuter != NULL)
		return CLASS_E_NOAGGREGATION;

	if (IsEqualIID(riid, IID_IShellExtInit) || IsEqualIID(riid, IID_IContextMenu)) {
		pHandler = new FileDaterContextMenuHandler();
		if (pHandler == NULL) {
			return E_OUTOFMEMORY;
		}
		hResult = pHandler->QueryInterface(riid, ppvObject);
		pHandler->Release();
	}
	else {
		hResult = E_NOINTERFACE;
	}

	return hResult;
}

HRESULT FileDaterClassFactory::LockServer(BOOL flock)
{
	return S_OK;
}

FileDaterClassFactory::~FileDaterClassFactory()
{
	InterlockedDecrement(&g_cObjCount);
}