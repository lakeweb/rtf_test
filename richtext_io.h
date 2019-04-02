#pragma once

#include <cstddef>

class _afxRichEditStreamCookie
{
public:
	CArchive& m_ar;
	DWORD m_dwError;
	_afxRichEditStreamCookie(CArchive& ar) : m_ar(ar) {m_dwError=0;}
};

using shared_char_ptr = std::unique_ptr<char>;

CString GetRichTextStr( CRichEditCtrl& re, int format= SF_RTF );
BOOL SetRichText( CRichEditCtrl& re, LPCTSTR buf, int format= SF_RTF );
BOOL SetRichText( CRichEditCtrl& re, CFile& file );
//format in the case of reading a txt file
BOOL SetRichText( CRichEditCtrl& re, std::istream& is, int format= SF_RTF );
shared_char_ptr GetRichText( CRichEditCtrl& re, int format= SF_RTF );
//Add a fully qualified RTF Block to the CRichEditCtrl
BOOL AppendRichText(CRichEditCtrl& re, LPCTSTR buf);

DWORD CALLBACK EditStreamCallBack(
	DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb );

/*
const char _RichEditPreamble[]= R"({\rtf1\ansi\deff0\deflang1033\
{\fonttbl{\f0\froman\fprq2\fcharset0 Times New Roman;}}\
{\colortbl\red0\green0\blue0;\red129\green129\blue129;\
\red128\green0\blue0;\red0\green128\blue0;}\
\viewkind4\uc1\pard\f0\fs24\sa120 )";

*/
