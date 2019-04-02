
#include "stdafx.h"
#include "richtext_io.h"


shared_char_ptr GetRichText( CRichEditCtrl& re, int format ) //| SFF_SELECTION
{
	CMemFile iFile( 1024 );
	CArchive oar( &iFile, CArchive::store );

	//need a throw exception here.
	EDITSTREAM es= { NULL, 0, EditStreamCallBack };
	_afxRichEditStreamCookie cookie( oar );
	es.dwCookie= (DWORD_PTR)&cookie;
	re.StreamOut( format, es );

	return shared_char_ptr((char*)iFile.Detach(),std::default_delete<char>());
}

DWORD CALLBACK EditStreamCallBack(
	DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb )
{
	_afxRichEditStreamCookie* pCookie = (_afxRichEditStreamCookie*)dwCookie;
	CArchive& ar = pCookie->m_ar;
	DWORD dw = 0;
	*pcb = cb;
	TRY
	{
		if ( ar.IsStoring( ) )
		ar.GetFile( )->Write( pbBuff, cb );
		else
			*pcb = ar.GetFile( )->Read( pbBuff, cb );
	}
		CATCH( CFileException, e )
	{
		*pcb = 0;
		pCookie->m_dwError = (DWORD)e->m_cause;
		dw = 1;
		e->Delete( );
	}
	AND_CATCH_ALL( e )
	{
		*pcb = 0;
		pCookie->m_dwError = -1;
		dw = 1;
		e->Delete( );
	}
	END_CATCH_ALL
		return dw;
}

BOOL SetRichText( CRichEditCtrl& re, LPCTSTR buf, int format )//| SFF_SELECTION
{
	CMemFile oFile( (BYTE*)buf, (UINT)_tcslen( buf ) );
	CArchive oar( &oFile, CArchive::load );

	EDITSTREAM es= {0, 0, EditStreamCallBack};
	_afxRichEditStreamCookie cookie( oar );
	es.dwCookie = (DWORD_PTR)&cookie;

	oFile.SeekToBegin( );
	re.StreamIn( format, es ); //| SFF_SELECTION
	if( cookie.m_dwError != 0 )
	{
		AfxThrowFileException( cookie.m_dwError );
		return FALSE;
	}

	oar.Close( );
	oFile.Close( );
	return TRUE;
}

BOOL SetRichText( CRichEditCtrl& re, CFile& file )//| SFF_SELECTION
{
	//CMemFile oFile( (BYTE*)buf, (UINT)strlen( buf ) );
	//CArchive oar( &file, CArchive::load, 4096 * 16 );
	CArchive oar( &file, CArchive::load, 64000 );

	EDITSTREAM es= {0, 0, EditStreamCallBack};
	_afxRichEditStreamCookie cookie( oar );
	es.dwCookie = (DWORD_PTR)&cookie;

	//oFile.Write( buf, strlen( buf ) );
	file.SeekToBegin( );
	re.StreamIn( SF_RTF, es ); //| SFF_SELECTION
	if( cookie.m_dwError != 0 )
	{
		AfxThrowFileException( cookie.m_dwError );
		return FALSE;
	}
	return TRUE;
}

BOOL AppendRichText( CRichEditCtrl& rtf, LPCTSTR buf )
{
	rtf.SetSel(-1, -1);
	if( ! SetRichText(rtf,buf, SF_RTF | SFF_SELECTION))
		return FALSE;
	auto buffer = GetRichText(rtf);
	char* che= buffer.get();
	for(; *che; ++che);//to end
	char* ch= che;
	for(; *ch != ' '; --ch);//back to first space
	for(; *ch != '\\'; ++ch);//then to first '\', assumes not \\,\},\{ for now
	if( ch + 10 > che )
		return FALSE;//but it should fit....
	auto re = R"(\par\par})"; // the replacement
	for( size_t i= 0; i < 10; ++i)
		*ch++ = *re++;
	return SetRichText(rtf,buffer.get());
}