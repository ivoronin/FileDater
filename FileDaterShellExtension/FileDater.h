#pragma once

class FileDater {
private:
	wchar_t m_wszSrcPathname[_MAX_PATH];
	HRESULT Stamp(BOOL rename);
public:
	wchar_t m_wszDstName[_MAX_FNAME];
	wchar_t m_wszDstDir[_MAX_DIR];
	wchar_t m_wszDstPathname[_MAX_PATH];
	FileDater(wchar_t * wszPathName);
	HRESULT Copy();
	HRESULT Rename();
	~FileDater();
};
