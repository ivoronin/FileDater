#pragma once
#include <unknwn.h>

class FileDaterClassFactory : public IClassFactory
{
protected:
	DWORD m_dwRefCount;
	~FileDaterClassFactory();
public:
	FileDaterClassFactory();

	/* IUnknown methods */
	ULONG __stdcall AddRef();
	ULONG __stdcall Release();
	HRESULT __stdcall QueryInterface(REFIID riid, void **ppvObject);

	/* IClassFactory methods */
	HRESULT __stdcall CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject);
	HRESULT __stdcall LockServer(BOOL flock);
};
