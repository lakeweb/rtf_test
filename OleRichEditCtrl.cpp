// OleRichEditCtrl.cpp : implementation file
//

#include "stdafx.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//...................................................
// COleRichEditCtrl

//#define ROLETRACE TRACE
#define ROLETRACE __noop
//

COleRichEditCtrl::COleRichEditCtrl( )
	:m_pIRichEditOleCallback( NULL )
	,m_hInstRichEdit50W2( NULL )
	,pIREOle( NULL )
{
	m_bCallbackSet = FALSE;
}

COleRichEditCtrl::~COleRichEditCtrl( )
{
	//CHandleMap* pMap= this->afxMapHWND( );

	//if( ! pMap )
	//{
#ifdef CONSOLE_APP
		::DestroyWindow( m_hWnd );
		m_hWnd= NULL;
#endif // CONSOLE_APP
	//}
	CRichEditCtrl::~CRichEditCtrl( );

	if( pIREOle )
		pIREOle->Release( );

	if( m_hInstRichEdit50W2 )
		FreeLibrary( m_hInstRichEdit50W2 );

	// IExRichEditOleCallback class is a reference-counted class  
	// which deletes itself and for which delete should not be called
}

BEGIN_MESSAGE_MAP( COleRichEditCtrl, CRichEditCtrl )
	ON_WM_CREATE( )
END_MESSAGE_MAP( )

BOOL COleRichEditCtrl::Create( DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID )
{
	//We have ond of these in 'RichView.cpp'
		m_hInstRichEdit50W2 = LoadLibrary( _T("msftedit.dll") );
	//m_hInstRichEdit50W2= NULL;

	if( !AfxInitRichEdit2( ) )
		return FALSE;

	if( m_hInstRichEdit50W2 )
	{
		m_hWnd=
			::CreateWindow( _T("RichEdit50W"),
			0, dwStyle, 0, 0, 0, 0, pParentWnd->GetSafeHwnd( ), NULL,
			AfxGetInstanceHandle( ), NULL );
	}
	else
	{
		m_hWnd=
			::CreateWindow( _T("RichEdit20W"), /*RICHEDIT_CLASS*/
			0, dwStyle, 0, 0, 0, 0, pParentWnd->GetSafeHwnd( ), NULL,
			AfxGetInstanceHandle( ), NULL );
	}
	if( ! m_hWnd )
		return FALSE;

	PreSubclassWindow( );
	if ( !m_bCallbackSet )
		SetOLECallback( m_pIRichEditOleCallback );

	//Must do this for tabs to work!!!!
	SendMessage( EM_SETTYPOGRAPHYOPTIONS, TO_ADVANCEDTYPOGRAPHY, TO_ADVANCEDTYPOGRAPHY );
	
	LRESULT rs= SendMessage( EM_GETOLEINTERFACE, 0, (LPARAM)&pIREOle );
	if( ! rs )
		pIREOle= NULL;

	return TRUE;
}

int COleRichEditCtrl::OnCreate( LPCREATESTRUCT lpCreateStruct ) 
{
 	if( CRichEditCtrl::OnCreate( lpCreateStruct ) == -1 )
 		return -1;

	// m_pIRichEditOleCallback should have been created in PreSubclassWindow
 	ASSERT( m_pIRichEditOleCallback != NULL );	

	// set the IExRichEditOleCallback pointer if it wasn't set 
	// successfully in PreSubclassWindow
	if ( !m_bCallbackSet )
	{
		SetOLECallback( m_pIRichEditOleCallback );
	}
 	 	return 0;
}

void COleRichEditCtrl::PreSubclassWindow( ) 
{
	//Only once and no one should be here as the only Create(..) does it
	ASSERT( ! m_pIRichEditOleCallback );
	// base class first
	CRichEditCtrl::PreSubclassWindow( );	

	m_pIRichEditOleCallback = new IExRichEditOleCallback;
	ASSERT( m_pIRichEditOleCallback != NULL );

	m_bCallbackSet = SetOLECallback( m_pIRichEditOleCallback );
}

COleRichEditCtrl::IExRichEditOleCallback::IExRichEditOleCallback()
{
	pStorage = NULL;
	m_iNumStorages = 0;
	m_dwRef = 0;

	// set up OLE storage

	HRESULT hResult = ::StgCreateDocfile(NULL,
		STGM_TRANSACTED | STGM_READWRITE | STGM_SHARE_EXCLUSIVE /*| STGM_DELETEONRELEASE */|STGM_CREATE ,
		0, &pStorage );

	if ( pStorage == NULL ||
		hResult != S_OK )
	{
		AfxThrowOleException( hResult );
	}
}

COleRichEditCtrl::IExRichEditOleCallback::~IExRichEditOleCallback()
{
}

HRESULT STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::GetNewStorage( LPSTORAGE* lplpstg )
{
	ROLETRACE( "GetNewStorage\n" );
	m_iNumStorages++;
	WCHAR tName[50];
	swprintf_s( tName, 50, L"REOLEStorage%d", m_iNumStorages );

	HRESULT hResult= pStorage->CreateStorage(
		tName, 
		STGM_TRANSACTED | STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE ,
		0, 0, lplpstg );

	if( hResult != S_OK )
	{
		::AfxThrowOleException( hResult );
	}
	return hResult;
}

HRESULT STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::QueryInterface(REFIID iid, void ** ppvObject)
{
	ROLETRACE( "QueryInterface\n" );
	HRESULT hr = S_OK;
	*ppvObject = NULL;
	
	if ( iid == IID_IUnknown ||
		iid == IID_IRichEditOleCallback )
	{
		*ppvObject = this;
		AddRef();
		hr = NOERROR;
	}
	else
	{
		hr = E_NOINTERFACE;
	}
	return hr;
}

ULONG STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::AddRef()
{
	return ++m_dwRef;
}

ULONG STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::Release()
{
	if ( --m_dwRef == 0 )
	{
		delete this;
		return 0;
	}
	return m_dwRef;
}

HRESULT STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::GetInPlaceContext(LPOLEINPLACEFRAME FAR *lplpFrame,
	LPOLEINPLACEUIWINDOW FAR *lplpDoc, LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
	ROLETRACE( "GetInPlaceContext\n" );
	return S_OK;
}

HRESULT STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::ShowContainerUI(BOOL fShow)
{
	ROLETRACE( "ShowContainerUI\n" );
	return S_OK;
}

HRESULT STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::QueryInsertObject(LPCLSID lpclsid, LPSTORAGE lpstg, LONG cp)
{
	if( lpclsid == (LPCLSID)0x0013F150 )
	{
		IStream *pStream;

		HRESULT res = pStorage->OpenStream( L"REOLEStorage1", NULL, STGM_SHARE_EXCLUSIVE | STGM_READ, 0, &pStream);
		if( res == 0 )
			++res;
	}
	ROLETRACE( "QueryInsertObject\n" );
	return S_OK;
}

HRESULT STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::DeleteObject(LPOLEOBJECT lpoleobj)
{
	ROLETRACE( "Delete RichEditOle Object\n" );
	return S_OK;
}

HRESULT STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::QueryAcceptData(LPDATAOBJECT lpdataobj, CLIPFORMAT FAR *lpcfFormat,
	DWORD reco, BOOL fReally, HGLOBAL hMetaPict)
{
	ROLETRACE( "QueryAcceptData odat: %x fmt: %x reco:%d really %d meta:%d \n", lpdataobj, *lpcfFormat, reco, fReally, hMetaPict );//, fDrag, grfKeyState, *pdwEffect );
	return S_OK;
}

HRESULT STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::ContextSensitiveHelp(BOOL fEnterMode)
{
	ROLETRACE( "ContextSensitiveHelp\n" );
	return S_OK;
}

HRESULT STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::GetClipboardData(CHARRANGE FAR *lpchrg, DWORD reco, LPDATAOBJECT FAR *lplpdataobj)
{
	ROLETRACE( "GetClipboardData\n" );
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::GetDragDropEffect(BOOL fDrag, DWORD grfKeyState, LPDWORD pdwEffect)
{
	ROLETRACE( "callback f:%d key: %d dw: %d\n", fDrag, grfKeyState, *pdwEffect );
	return S_OK;
}

HRESULT STDMETHODCALLTYPE 
COleRichEditCtrl::IExRichEditOleCallback::GetContextMenu(WORD seltyp, LPOLEOBJECT lpoleobj, CHARRANGE FAR *lpchrg,
	HMENU FAR *lphmenu)
{
	ROLETRACE( "GetContextMenu\n" );
	return S_OK;
}


//We use the global (see globals.h) so don't need this
//long COleRichEditCtrl::StreamInFromResource(int iRes, LPCTSTR sType)
//{
//	HINSTANCE hInst = AfxGetInstanceHandle();
//	HRSRC hRsrc = ::FindResource(hInst,
//		MAKEINTRESOURCE(iRes), sType);
//	
//	DWORD len = SizeofResource(hInst, hRsrc); 
//	BYTE* lpRsrc = (BYTE*)LoadResource(hInst, hRsrc); 
//	ASSERT(lpRsrc); 
// 
//	CMemFile mfile;
//	mfile.Attach(lpRsrc, len); 
//
//	EDITSTREAM es;
//	es.pfnCallback = readFunction;
//	es.dwError = 0;
//	es.dwCookie = (DWORD_PTR) &mfile;
//
//	return StreamIn( SF_RTF, es );
//}

/* static */
//DWORD CALLBACK COleRichEditCtrl::readFunction(DWORD dwCookie,
//		 LPBYTE lpBuf,			// the buffer to fill
//		 LONG nCount,			// number of bytes to read
//		 LONG* nRead)			// number of bytes actually read
//{
//	CFile* fp = (CFile *)(DWORD_PTR)dwCookie;
//	*nRead = fp->Read(lpBuf,nCount);
//	return 0;
//}


