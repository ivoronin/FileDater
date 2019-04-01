#include "stdafx.h"
#include <stdexcept>
#include <regex>
#include <strsafe.h>
#include <Shobjidl.h>
#include "FileDater.h"

UINT ParseFilename(const wchar_t *fname, wchar_t *bname, size_t bsize, wchar_t *date, size_t dsize, UINT *version) {
	std::wregex pattern(L"^([^\\.]+?)(?:\\s(\\d{6})-(\\d{2}))$");
	std::wcmatch results;

	if (nullptr == fname || nullptr == bname || nullptr == date || nullptr == version)
		throw std::invalid_argument("nullptr passed");

	if (std::regex_search(fname, results, pattern)) {
		if (wcscpy_s(bname, bsize, results[1].str().c_str()) != 0)
			throw std::runtime_error("wcscpy_s(bname, results) error");
		if (wcscpy_s(date, dsize, results[2].str().c_str()) != 0)
			throw std::runtime_error("wcscpy_s(date, results) error");
		*version = _wtoi(results[3].str().c_str());
		if (errno != NOERROR)
			throw std::runtime_error("_wtoi(results) error");
		return 0;
	}
	return 1;
}

void PutCurrentDate(wchar_t * date, size_t dsize) {
	time_t tTime;
	struct tm tmTime;

	tTime = time(nullptr);
	if (tTime == -1)
		throw std::runtime_error("time() error");
	if (localtime_s(&tmTime, &tTime) != 0)
		throw std::runtime_error("localtime_s() error");
	if (wcsftime(date, dsize, L"%d%m%y", &tmTime) == 0)
		throw std::runtime_error("wcsftime() error");
}

#define _MAX_DATE 16
FileDater::FileDater(wchar_t * szPathname)
{
	wchar_t drive[_MAX_DRIVE], dir[_MAX_DIR], fname[_MAX_FNAME], ext[_MAX_EXT];
	wchar_t bname[_MAX_PATH], date[_MAX_DATE];
	UINT version;
	wchar_t new_date[_MAX_DATE] = L"";
	wchar_t new_fname[_MAX_PATH];
	UINT new_version = 1;

	if (wcscpy_s(m_szSrcPathname, _MAX_PATH, szPathname) != 0)
		throw std::runtime_error("wcscpy_s(m_szSrcPathname,szPathname) error");

	OutputDebugString(L"HERE");
	OutputDebugString(m_szSrcPathname);

	if (_wsplitpath_s(m_szSrcPathname, drive, _countof(drive), dir, _countof(dir), fname, _countof(fname), ext, _countof(ext)) != 0)
		throw std::runtime_error("_wsplitpath_s(szPathname) error");

	OutputDebugString(fname);

	if (wcslen(fname) == 0)
		throw std::runtime_error("fname is empty");

	PutCurrentDate(new_date, _MAX_DATE);

	if (ParseFilename(fname, bname, _MAX_PATH, date, _MAX_DATE, &version) == 0) {
		if (wcscmp(date, new_date) == 0)
			new_version = version + 1;
	}
	else {
		if (wcscpy_s(bname, _MAX_PATH, fname) != 0)
			throw std::runtime_error("wcscpy_s(bname, fname) error");
	}

	OutputDebugString(bname);

	StringCbPrintf(new_fname, _MAX_FNAME, L"%s %s-%02d", bname, new_date, new_version);

	OutputDebugString(new_fname);

	if(_wmakepath_s(m_szDstName, _MAX_FNAME, NULL, NULL, new_fname, ext) != 0)
		throw std::runtime_error("_wmakepath_s(m_szDstName) error");
	if(_wmakepath_s(m_szDstDir, _MAX_DIR, drive, dir, NULL, NULL) != 0)
		throw std::runtime_error("_wmakepath_s(m_szDstDir) error");
	if(_wmakepath_s(m_szDstPathname, _MAX_PATH, drive, dir, new_fname, ext) != 0)
		throw std::runtime_error("_wmakepath_s(m_szDstPathname) error");
}

HRESULT FileDater::Stamp(BOOL rename)
{
	HRESULT hr;
	IFileOperation *pfo;
	IShellItem *psiSrc = NULL;
	IShellItem *psiDstDir = NULL;

	hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
	if (SUCCEEDED(hr)) {
		hr = CoCreateInstance(CLSID_FileOperation, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pfo));
		if (SUCCEEDED(hr)) {
			hr = SHCreateItemFromParsingName(m_szSrcPathname, NULL, IID_PPV_ARGS(&psiSrc));
			if (SUCCEEDED(hr)) {
				hr = SHCreateItemFromParsingName(m_szDstDir, NULL, IID_PPV_ARGS(&psiDstDir));
				if (SUCCEEDED(hr)) {
					if (rename)
						hr = pfo->MoveItem(psiSrc, psiDstDir, m_szDstName, NULL);
					else
						hr = pfo->CopyItem(psiSrc, psiDstDir, m_szDstName, NULL);
					if (SUCCEEDED(hr)) {
						hr = pfo->PerformOperations();
					}
					psiDstDir->Release();
				}
				psiSrc->Release();
			}
			pfo->Release();
		}
		CoUninitialize();
	}

	return hr;
}

HRESULT FileDater::Rename()
{
	return Stamp(TRUE);
}

HRESULT FileDater::Copy()
{
	return Stamp(FALSE);
}

FileDater::~FileDater()
{
}
