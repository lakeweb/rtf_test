#if !defined( OLERICHEDITCTRL_H )
#define OLERICHEDITCTRL_H

#ifdef COleRichEditCtrl
#undef COleRichEditCtrl
#endif

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// OleRichEditCtrl.h : header file
//

//this is in the 7.1 SDK TODO Update SDK????
#include <richole.h>

/////////////////////////////////////////////////////////////////////////////
// COleRichEditCtrl window

class COleRichEditCtrl : public CRichEditCtrl
{
	HINSTANCE m_hInstRichEdit50W2; // handle to MSFTEDIT.DLL

	//new 2/18/10
	LPRICHEDITOLE pIREOle;

public:
	COleRichEditCtrl( );
	virtual ~COleRichEditCtrl( );

	long StreamInFromResource( int iRes, LPCTSTR sType );
	BOOL Create( ) { return Create( ES_MULTILINE, CRect( 0, 0, 0, 0 ), NULL, 0 ); }
	BOOL Create( DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID );
	size_t Print( CDC* pDC, LPRECT pRect );
	size_t PrintHEx( CDC* pDC, LPRECT pRect );

	HRESULT InsertObject( REOBJECT* pObj ) { return pIREOle->InsertObject( pObj ); }
	HRESULT GetClientSite( LPOLECLIENTSITE* pCS ) { return pIREOle->GetClientSite( pCS ); }
	HRESULT ReleaseOle( ) { return pIREOle->Release( ); }

	//testing
//#ifdef _DEBUG
	LPRICHEDITOLE GetRichOle( ) { return pIREOle; }
//#endif

protected:
	virtual void PreSubclassWindow( );

protected:
	DECLARE_MESSAGE_MAP( )
	afx_msg int OnCreate( LPCREATESTRUCT lpCreateStruct );

protected:
	static DWORD CALLBACK readFunction( DWORD dwCookie,
		 LPBYTE lpBuf,			// the buffer to fill
		 LONG nCount,			// number of bytes to read
		 LONG* nRead );			// number of bytes actually read

	interface IExRichEditOleCallback;	// forward declaration (see below in this header file)

	IExRichEditOleCallback* m_pIRichEditOleCallback;
	BOOL m_bCallbackSet;

	interface IExRichEditOleCallback : public IRichEditOleCallback
	{
	public:
		IExRichEditOleCallback( );
		virtual ~IExRichEditOleCallback( );
		int m_iNumStorages;
		IStorage* pStorage;
		DWORD m_dwRef;

		virtual HRESULT STDMETHODCALLTYPE GetNewStorage(LPSTORAGE* lplpstg);
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void ** ppvObject);
		virtual ULONG STDMETHODCALLTYPE AddRef();
		virtual ULONG STDMETHODCALLTYPE Release();
		virtual HRESULT STDMETHODCALLTYPE GetInPlaceContext(LPOLEINPLACEFRAME FAR *lplpFrame,
			LPOLEINPLACEUIWINDOW FAR *lplpDoc, LPOLEINPLACEFRAMEINFO lpFrameInfo);
 		virtual HRESULT STDMETHODCALLTYPE ShowContainerUI(BOOL fShow);
 		virtual HRESULT STDMETHODCALLTYPE QueryInsertObject(LPCLSID lpclsid, LPSTORAGE lpstg, LONG cp);
 		virtual HRESULT STDMETHODCALLTYPE DeleteObject(LPOLEOBJECT lpoleobj);
 		virtual HRESULT STDMETHODCALLTYPE QueryAcceptData(LPDATAOBJECT lpdataobj, CLIPFORMAT FAR *lpcfFormat,
			DWORD reco, BOOL fReally, HGLOBAL hMetaPict);
 		virtual HRESULT STDMETHODCALLTYPE ContextSensitiveHelp(BOOL fEnterMode);
 		virtual HRESULT STDMETHODCALLTYPE GetClipboardData(CHARRANGE FAR *lpchrg, DWORD reco, LPDATAOBJECT FAR *lplpdataobj);
 		virtual HRESULT STDMETHODCALLTYPE GetDragDropEffect(BOOL fDrag, DWORD grfKeyState, LPDWORD pdwEffect);
 		virtual HRESULT STDMETHODCALLTYPE GetContextMenu(WORD seltyp, LPOLEOBJECT lpoleobj, CHARRANGE FAR *lpchrg,
			HMENU FAR *lphmenu);
	};
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined( OLERICHEDITCTRL_H )
