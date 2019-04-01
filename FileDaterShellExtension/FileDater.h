#pragma once

class FileDater {
private:
	wchar_t m_szSrcPathname[_MAX_PATH];
	HRESULT Stamp(BOOL rename);
public:
	wchar_t m_szDstName[_MAX_FNAME];
	wchar_t m_szDstDir[_MAX_DIR];
	wchar_t m_szDstPathname[_MAX_PATH];
	FileDater(wchar_t * szPathName);
	HRESULT Copy();
	HRESULT Rename();
	~FileDater();
};
