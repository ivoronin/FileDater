#include "stdafx.h"
#include <stdexcept>
#include <regex>
#include <strsafe.h>
#include <Shobjidl.h>
#include "FileDater.h"

UINT ParseFilename(const wchar_t *wszFname, wchar_t *wszBname, size_t cBnameSize, wchar_t *wszDate, size_t cDateSize, UINT *iVersion) {
	std::wregex pattern(L"^([^\\.]+?)(?:\\s(\\d{6})-(\\d{2}))$");
	std::wcmatch results;

	if (nullptr == wszFname || nullptr == wszBname || nullptr == wszDate || nullptr == iVersion)
		throw std::invalid_argument("nullptr passed");

	if (std::regex_search(wszFname, results, pattern)) {
		if (wcscpy_s(wszBname, cBnameSize, results[1].str().c_str()) != 0)
			throw std::runtime_error("wcscpy_s(wszBname, results) error");
		if (wcscpy_s(wszDate, cDateSize, results[2].str().c_str()) != 0)
			throw std::runtime_error("wcscpy_s(wszDate, results) error");
		*iVersion = _wtoi(results[3].str().c_str());
		if (errno != NOERROR)
			throw std::runtime_error("_wtoi(results) error");
		return 0;
	}
	return 1;
}

void PutCurrentDate(wchar_t * wszDate, size_t cDateSize) {
	time_t tTime;
	struct tm tmTime;

	tTime = time(nullptr);
	if (tTime == -1)
		throw std::runtime_error("time() error");
	if (localtime_s(&tmTime, &tTime) != 0)
		throw std::runtime_error("localtime_s() error");
	if (wcsftime(wszDate, cDateSize, L"%d%m%y", &tmTime) == 0)
		throw std::runtime_error("wcsftime() error");
}

#define _MAX_DATE 16
FileDater::FileDater(wchar_t * wszPathname)
{
	wchar_t wszDrive[_MAX_DRIVE], wszDir[_MAX_DIR], wszFname[_MAX_FNAME], wszExt[_MAX_EXT];
	wchar_t wszBname[_MAX_PATH], wszDate[_MAX_DATE];
	UINT iVersion;
	wchar_t wszNewDate[_MAX_DATE] = L"";
	wchar_t wszNewFname[_MAX_PATH];
	UINT iNewVersion = 1;

	if (wcscpy_s(m_wszSrcPathname, _MAX_PATH, wszPathname) != 0)
		throw std::runtime_error("wcscpy_s(m_wszSrcPathname, wszPathname) error");

	if (_wsplitpath_s(m_wszSrcPathname, wszDrive, _countof(wszDrive), wszDir, _countof(wszDir), wszFname, _countof(wszFname), wszExt, _countof(wszExt)) != 0)
		throw std::runtime_error("_wsplitpath_s(wszPathname) error");

	if (wcslen(wszFname) == 0)
		throw std::runtime_error("wszFname is empty");

	PutCurrentDate(wszNewDate, _MAX_DATE);

	if (ParseFilename(wszFname, wszBname, _MAX_PATH, wszDate, _MAX_DATE, &iVersion) == 0) {
		if (wcscmp(wszDate, wszNewDate) == 0)
			iNewVersion = iVersion + 1;
	}
	else {
		if (wcscpy_s(wszBname, _MAX_PATH, wszFname) != 0)
			throw std::runtime_error("wcscpy_s(wszBname, wszFname) error");
	}

	StringCbPrintf(wszNewFname, _MAX_FNAME, L"%s %s-%02d", wszBname, wszNewDate, iNewVersion);

	if(_wmakepath_s(m_wszDstName, _MAX_FNAME, NULL, NULL, wszNewFname, wszExt) != 0)
		throw std::runtime_error("_wmakepath_s(m_wszDstName) error");
	if(_wmakepath_s(m_wszDstDir, _MAX_DIR, wszDrive, wszDir, NULL, NULL) != 0)
		throw std::runtime_error("_wmakepath_s(m_wszDstDir) error");
	if(_wmakepath_s(m_wszDstPathname, _MAX_PATH, wszDrive, wszDir, wszNewFname, wszExt) != 0)
		throw std::runtime_error("_wmakepath_s(m_wszDstPathname) error");
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
			hr = SHCreateItemFromParsingName(m_wszSrcPathname, NULL, IID_PPV_ARGS(&psiSrc));
			if (SUCCEEDED(hr)) {
				hr = SHCreateItemFromParsingName(m_wszDstDir, NULL, IID_PPV_ARGS(&psiDstDir));
				if (SUCCEEDED(hr)) {
					if (rename)
						hr = pfo->MoveItem(psiSrc, psiDstDir, m_wszDstName, NULL);
					else
						hr = pfo->CopyItem(psiSrc, psiDstDir, m_wszDstName, NULL);
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
