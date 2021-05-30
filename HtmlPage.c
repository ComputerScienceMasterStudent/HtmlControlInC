#include <windows.h>
#include <exdisp.h>		
#include <mshtml.h>		
#include <ExDispid.h>
#include <initguid.h> 


DEFINE_GUID(CLSID_IEventHandler, 0xb5b3d8f, 0x574c, 0x4fa3,
	0x90, 0x10, 0x25, 0xb8, 0xe4, 0xce, 0x24, 0xc2);

DEFINE_GUID(IID_IEventHandler, 0x74666cad, 0xc2b1, 0x4fa8,
	0xa0, 0x49, 0x97, 0xf3, 0x21, 0x48, 0x2, 0xf0);


#undef  INTERFACE
#define INTERFACE IEventHandler
DECLARE_INTERFACE_(INTERFACE, DWebBrowserEvents2)
{
	// IUnkown function
	STDMETHOD(QueryInterface)(THIS_ REFIID, void**) PURE;
	STDMETHOD_(ULONG, AddRef)(THIS) PURE;
	STDMETHOD_(ULONG, Release)(THIS) PURE;
	// IDispatch function
	STDMETHOD_(ULONG, GetTypeInfoCount)(THIS_ UINT*) PURE;
	STDMETHOD_(ULONG, GetTypeInfo)(THIS_ UINT, LCID, ITypeInfo**) PURE;
	STDMETHOD_(ULONG, GetIDsOfNames)(THIS_ REFIID, LPOLESTR*, UINT,
		LCID, DISPID*) PURE;
	STDMETHOD_(ULONG, Invoke)(THIS_ DISPID, REFIID, LCID, WORD,
		DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) PURE;
};

typedef struct {
	IEventHandlerVtbl* lpVtbl;
	DWORD			   count;
} MyRealEventHandler;


HRESULT STDMETHODCALLTYPE EvtQueryInterface(void* this,
	REFIID vTableGuid, void** ppv)
{
	if (!IsEqualIID(vTableGuid, &IID_IEventHandler) && !IsEqualIID(vTableGuid, &DIID_DWebBrowserEvents2))
	{
		*ppv = 0;
		return(E_NOINTERFACE);
	}
	*ppv = this;
	((MyRealEventHandler*)this)->lpVtbl->AddRef(this);

	return(NOERROR);
}

ULONG STDMETHODCALLTYPE EvtAddRef(void* this)
{
	InterlockedIncrement(&(((MyRealEventHandler*)this)->count));
	return(((MyRealEventHandler*)this)->count);
}

ULONG STDMETHODCALLTYPE EvtRelease(void* this)
{
	InterlockedIncrement(&(((MyRealEventHandler*)this)->count));
	if (((MyRealEventHandler*)this)->count == 0)
	{
		GlobalFree(this);
		return(0);
	}
	return(((MyRealEventHandler*)this)->count);
}
HRESULT STDMETHODCALLTYPE EvtGetTypeInfoCount(void* this, unsigned int* pctInfo)
{
	if (pctInfo == NULL) {
		return E_INVALIDARG;
	}
	*pctInfo = 1;
	return NOERROR;
}

HRESULT STDMETHODCALLTYPE EvtGetTypeInfo(void* this, unsigned int iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
	*ppTInfo = 0;
	return E_NOTIMPL;
}

STDMETHODIMP STDMETHODCALLTYPE EvtGetIDsOfNames(void* this, REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	UNREFERENCED_PARAMETER(rgszNames);
	UNREFERENCED_PARAMETER(cNames);
	UNREFERENCED_PARAMETER(lcid);
	UNREFERENCED_PARAMETER(rgDispId);

	return E_NOTIMPL;
}


HRESULT STDMETHODCALLTYPE EvtInvoke(void* this,
	DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	UNREFERENCED_PARAMETER(lcid);
	UNREFERENCED_PARAMETER(wFlags);
	UNREFERENCED_PARAMETER(pVarResult);
	UNREFERENCED_PARAMETER(pExcepInfo);
	UNREFERENCED_PARAMETER(puArgErr);
	if (!IsEqualIID(riid, &IID_NULL)) 
		return DISP_E_UNKNOWNINTERFACE; 

	BSTR HtmlURL;
	if (dispIdMember == DISPID_NEWWINDOW2) {
	}
	if (dispIdMember == DISPID_BEFORENAVIGATE2) { 
		IWebBrowser2* pSite = (IWebBrowser2*)pDispParams->rgvarg[6].pdispVal;
		pSite->lpVtbl->get_LocationURL(pSite, &HtmlURL);
	}
	return S_OK;
}

IEventHandlerVtbl       myEventHandlerTable = { EvtQueryInterface,EvtAddRef,EvtRelease,EvtGetTypeInfoCount,EvtGetTypeInfo,EvtGetIDsOfNames,EvtInvoke};
static DWORD		    objectsCount;
static DWORD		    serverLockCount;
static IClassFactory	myFactory;
static ULONG STDMETHODCALLTYPE FactoryAddRef(IClassFactory* this)
{
	InterlockedIncrement(&objectsCount);
	return(1);
}
static HRESULT STDMETHODCALLTYPE FactoryQueryInterface(IClassFactory* this, REFIID factoryGuid, void** ppv)
{
	if (IsEqualIID(factoryGuid, &IID_IEventHandler) || IsEqualIID(factoryGuid, &IID_IClassFactory))
	{
		this->lpVtbl->AddRef(this);
		return(NOERROR);
	}
	return(E_NOINTERFACE);
}
static ULONG STDMETHODCALLTYPE FactoryRelease(IClassFactory* this)
{
	return(InterlockedDecrement(&objectsCount));
}
static HRESULT STDMETHODCALLTYPE FactoryCreateInstance(IClassFactory* this, IUnknown* punkOuter, REFIID vTableGuid, void** objHandle)
{
	HRESULT				hr;
	register IEventHandler* thisobj;
	*objHandle = 0;
	if (punkOuter)
		hr = CLASS_E_NOAGGREGATION;
	else
	{
		if (!(thisobj = (IEventHandler*)GlobalAlloc(GMEM_FIXED, sizeof(MyRealEventHandler))))
			hr = E_OUTOFMEMORY;
		else
		{
			thisobj->lpVtbl = (IEventHandlerVtbl*)&myEventHandlerTable;
				((MyRealEventHandler*)thisobj)->count = 1;
				hr = myEventHandlerTable.QueryInterface(thisobj, vTableGuid, objHandle);
			myEventHandlerTable.Release(thisobj);
			if (SUCCEEDED(hr)) 
				InterlockedIncrement(&objectsCount);
		}
	}

	return(hr);
}

static HRESULT STDMETHODCALLTYPE FactoryLockServer(IClassFactory* this, BOOL flock)
{
	if (flock) InterlockedIncrement(&serverLockCount);
	else InterlockedDecrement(&serverLockCount);

	return(NOERROR);
}

static const IClassFactoryVtbl IClassFactory_Vtbl = { FactoryQueryInterface,
FactoryAddRef,
FactoryRelease,
FactoryCreateInstance,
FactoryLockServer };


unsigned char wndCount = 0;
static const TCHAR	className[] = "HtmlExample";

HRESULT STDMETHODCALLTYPE StorageQueryInterface(IStorage FAR* This, REFIID riid, LPVOID FAR* ppvObj);
HRESULT STDMETHODCALLTYPE StorageAddRef(IStorage FAR* This);
HRESULT STDMETHODCALLTYPE StorageRelease(IStorage FAR* This);
HRESULT STDMETHODCALLTYPE StorageCreateStream(IStorage FAR* This, const WCHAR *pwcsName, DWORD grfMode, DWORD reserved1, DWORD reserved2, IStream **ppstm);
HRESULT STDMETHODCALLTYPE StorageOpenStream(IStorage FAR* This, const WCHAR * pwcsName, void *reserved1, DWORD grfMode, DWORD reserved2, IStream **ppstm);
HRESULT STDMETHODCALLTYPE StorageCreateStorage(IStorage FAR* This, const WCHAR *pwcsName, DWORD grfMode, DWORD reserved1, DWORD reserved2, IStorage **ppstg);
HRESULT STDMETHODCALLTYPE StorageOpenStorage(IStorage FAR* This, const WCHAR * pwcsName, IStorage * pstgPriority, DWORD grfMode, SNB snbExclude, DWORD reserved, IStorage **ppstg);
HRESULT STDMETHODCALLTYPE StorageCopyTo(IStorage FAR* This, DWORD ciidExclude, IID const *rgiidExclude, SNB snbExclude,IStorage *pstgDest);
HRESULT STDMETHODCALLTYPE StorageMoveElementTo(IStorage FAR* This, const OLECHAR *pwcsName,IStorage * pstgDest, const OLECHAR *pwcsNewName, DWORD grfFlags);
HRESULT STDMETHODCALLTYPE StorageCommit(IStorage FAR* This, DWORD grfCommitFlags);
HRESULT STDMETHODCALLTYPE StorageRevert(IStorage FAR* This);
HRESULT STDMETHODCALLTYPE StorageEnumElements(IStorage FAR* This, DWORD reserved1, void * reserved2, DWORD reserved3, IEnumSTATSTG ** ppenum);
HRESULT STDMETHODCALLTYPE StorageDestroyElement(IStorage FAR* This, const OLECHAR *pwcsName);
HRESULT STDMETHODCALLTYPE StorageRenameElement(IStorage FAR* This, const WCHAR *pwcsOldName, const WCHAR *pwcsNewName);
HRESULT STDMETHODCALLTYPE StorageSetElementTimes(IStorage FAR* This, const WCHAR *pwcsName, FILETIME const *pctime, FILETIME const *patime, FILETIME const *pmtime);
HRESULT STDMETHODCALLTYPE StorageSetClass(IStorage FAR* This, REFCLSID clsid);
HRESULT STDMETHODCALLTYPE StorageSetStateBits(IStorage FAR* This, DWORD grfStateBits, DWORD grfMask);
HRESULT STDMETHODCALLTYPE StorageStat(IStorage FAR* This, STATSTG * pstatstg, DWORD grfStatFlag);

IStorageVtbl MyIStorageTable = {StorageQueryInterface,
StorageAddRef,
StorageRelease,
StorageCreateStream,
StorageOpenStream,
StorageCreateStorage,
StorageOpenStorage,
StorageCopyTo,
StorageMoveElementTo,
StorageCommit,
StorageRevert,
StorageEnumElements,
StorageDestroyElement,
StorageRenameElement,
StorageSetElementTimes,
StorageSetClass,
StorageSetStateBits,
StorageStat};

IStorage			MyIStorage = { &MyIStorageTable };

HRESULT STDMETHODCALLTYPE FrameQueryInterface(IOleInPlaceFrame FAR* This, REFIID riid, LPVOID FAR* ppvObj);
HRESULT STDMETHODCALLTYPE FrameAddRef(IOleInPlaceFrame FAR* This);
HRESULT STDMETHODCALLTYPE FrameRelease(IOleInPlaceFrame FAR* This);
HRESULT STDMETHODCALLTYPE FrameGetWindow(IOleInPlaceFrame FAR* This, HWND FAR* lphwnd);
HRESULT STDMETHODCALLTYPE FrameContextSensitiveHelp(IOleInPlaceFrame FAR* This, BOOL fEnterMode);
HRESULT STDMETHODCALLTYPE FrameGetBorder(IOleInPlaceFrame FAR* This, LPRECT lprectBorder);
HRESULT STDMETHODCALLTYPE FrameRequestBorderSpace(IOleInPlaceFrame FAR* This, LPCBORDERWIDTHS pborderwidths);
HRESULT STDMETHODCALLTYPE FrameSetBorderSpace(IOleInPlaceFrame FAR* This, LPCBORDERWIDTHS pborderwidths);
HRESULT STDMETHODCALLTYPE FrameSetActiveObject(IOleInPlaceFrame FAR* This, IOleInPlaceActiveObject *pActiveObject, LPCOLESTR pszObjName);
HRESULT STDMETHODCALLTYPE FrameInsertMenus(IOleInPlaceFrame FAR* This, HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths);
HRESULT STDMETHODCALLTYPE FrameSetMenu(IOleInPlaceFrame FAR* This, HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject);
HRESULT STDMETHODCALLTYPE FrameRemoveMenus(IOleInPlaceFrame FAR* This, HMENU hmenuShared);
HRESULT STDMETHODCALLTYPE FrameSetStatusText(IOleInPlaceFrame FAR* This, LPCOLESTR pszStatusText);
HRESULT STDMETHODCALLTYPE FrameEnableModeless(IOleInPlaceFrame FAR* This, BOOL fEnable);
HRESULT STDMETHODCALLTYPE FrameTranslateAccelerator(IOleInPlaceFrame FAR* This, LPMSG lpmsg, WORD wID);

IOleInPlaceFrameVtbl MyIOleInPlaceFrameTable = {FrameQueryInterface,
FrameAddRef,
FrameRelease,
FrameGetWindow,
FrameContextSensitiveHelp,
FrameGetBorder,
FrameRequestBorderSpace,
FrameSetBorderSpace,
FrameSetActiveObject,
FrameInsertMenus,
FrameSetMenu,
FrameRemoveMenus,
FrameSetStatusText,
FrameEnableModeless,
FrameTranslateAccelerator};

typedef struct _IOleInPlaceFrameEx {
	IOleInPlaceFrame	frame;		
	HWND				window;
} IOleInPlaceFrameEx;


HRESULT STDMETHODCALLTYPE SiteQueryInterface(IOleClientSite FAR* This, REFIID riid, void ** ppvObject);
HRESULT STDMETHODCALLTYPE SiteAddRef(IOleClientSite FAR* This);
HRESULT STDMETHODCALLTYPE SiteRelease(IOleClientSite FAR* This);
HRESULT STDMETHODCALLTYPE SiteSaveObject(IOleClientSite FAR* This);
HRESULT STDMETHODCALLTYPE SiteGetMoniker(IOleClientSite FAR* This, DWORD dwAssign, DWORD dwWhichMoniker, IMoniker ** ppmk);
HRESULT STDMETHODCALLTYPE SiteGetContainer(IOleClientSite FAR* This, LPOLECONTAINER FAR* ppContainer);
HRESULT STDMETHODCALLTYPE SiteShowObject(IOleClientSite FAR* This);
HRESULT STDMETHODCALLTYPE SiteOnShowWindow(IOleClientSite FAR* This, BOOL fShow);
HRESULT STDMETHODCALLTYPE SiteRequestNewObjectLayout(IOleClientSite FAR* This);

IOleClientSiteVtbl MyIOleClientSiteTable = {SiteQueryInterface,
SiteAddRef,
SiteRelease,
SiteSaveObject,
SiteGetMoniker,
SiteGetContainer,
SiteShowObject,
SiteOnShowWindow,
SiteRequestNewObjectLayout};



HRESULT STDMETHODCALLTYPE SiteGetWindow(IOleInPlaceSite FAR* This, HWND FAR* lphwnd);
HRESULT STDMETHODCALLTYPE SiteContextSensitiveHelp(IOleInPlaceSite FAR* This, BOOL fEnterMode);
HRESULT STDMETHODCALLTYPE SiteCanInPlaceActivate(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE SiteOnInPlaceActivate(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE SiteOnUIActivate(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE SiteGetWindowContext(IOleInPlaceSite FAR* This, LPOLEINPLACEFRAME FAR* lplpFrame,LPOLEINPLACEUIWINDOW FAR* lplpDoc,LPRECT lprcPosRect,LPRECT lprcClipRect,LPOLEINPLACEFRAMEINFO lpFrameInfo);
HRESULT STDMETHODCALLTYPE SiteScroll(IOleInPlaceSite FAR* This, SIZE scrollExtent);
HRESULT STDMETHODCALLTYPE SiteOnUIDeactivate(IOleInPlaceSite FAR* This, BOOL fUndoable);
HRESULT STDMETHODCALLTYPE SiteOnInPlaceDeactivate(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE SiteDiscardUndoState(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE SiteDeactivateAndUndo(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE SiteOnPosRectChange(IOleInPlaceSite FAR* This, LPCRECT lprcPosRect);

IOleInPlaceSiteVtbl MyIOleInPlaceSiteTable =  {SiteQueryInterface,
SiteAddRef,				
SiteRelease,				
SiteGetWindow,
SiteContextSensitiveHelp,
SiteCanInPlaceActivate,
SiteOnInPlaceActivate,
SiteOnUIActivate,
SiteGetWindowContext,
SiteScroll,
SiteOnUIDeactivate,
SiteOnInPlaceDeactivate,
SiteDiscardUndoState,
SiteDeactivateAndUndo,
SiteOnPosRectChange};

typedef struct __IOleInPlaceSiteEx {
	IOleInPlaceSite		inplace;		
	IOleInPlaceFrameEx	*frame;
} _IOleInPlaceSiteEx;

typedef struct __IOleClientSiteEx {
	IOleClientSite		client;			
	_IOleInPlaceSiteEx	inplace;		
} _IOleClientSiteEx;




HRESULT STDMETHODCALLTYPE StorageQueryInterface(IStorage FAR* This, REFIID riid, LPVOID FAR* ppvObj)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageAddRef(IStorage FAR* This)
{
	return(1);
}

HRESULT STDMETHODCALLTYPE StorageRelease(IStorage FAR* This)
{
	return(1);
}

HRESULT STDMETHODCALLTYPE StorageCreateStream(IStorage FAR* This, const WCHAR *pwcsName, DWORD grfMode, DWORD reserved1, DWORD reserved2, IStream **ppstm)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageOpenStream(IStorage FAR* This, const WCHAR * pwcsName, void *reserved1, DWORD grfMode, DWORD reserved2, IStream **ppstm)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageCreateStorage(IStorage FAR* This, const WCHAR *pwcsName, DWORD grfMode, DWORD reserved1, DWORD reserved2, IStorage **ppstg)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageOpenStorage(IStorage FAR* This, const WCHAR * pwcsName, IStorage * pstgPriority, DWORD grfMode, SNB snbExclude, DWORD reserved, IStorage **ppstg)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageCopyTo(IStorage FAR* This, DWORD ciidExclude, IID const *rgiidExclude, SNB snbExclude,IStorage *pstgDest)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageMoveElementTo(IStorage FAR* This, const OLECHAR *pwcsName,IStorage * pstgDest, const OLECHAR *pwcsNewName, DWORD grfFlags)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageCommit(IStorage FAR* This, DWORD grfCommitFlags)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageRevert(IStorage FAR* This)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageEnumElements(IStorage FAR* This, DWORD reserved1, void * reserved2, DWORD reserved3, IEnumSTATSTG ** ppenum)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageDestroyElement(IStorage FAR* This, const OLECHAR *pwcsName)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageRenameElement(IStorage FAR* This, const WCHAR *pwcsOldName, const WCHAR *pwcsNewName)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageSetElementTimes(IStorage FAR* This, const WCHAR *pwcsName, FILETIME const *pctime, FILETIME const *patime, FILETIME const *pmtime)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageSetClass(IStorage FAR* This, REFCLSID clsid)
{
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE StorageSetStateBits(IStorage FAR* This, DWORD grfStateBits, DWORD grfMask)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE StorageStat(IStorage FAR* This, STATSTG * pstatstg, DWORD grfStatFlag)
{
	return(E_NOTIMPL);
}


HRESULT STDMETHODCALLTYPE SiteQueryInterface(IOleClientSite FAR* This, REFIID riid, void ** ppvObject)
{
	if (!memcmp(riid, &IID_IUnknown, sizeof(GUID)) || !memcmp(riid, &IID_IOleClientSite, sizeof(GUID)))
		*ppvObject = &((_IOleClientSiteEx *)This)->client;

	else if (!memcmp(riid, &IID_IOleInPlaceSite, sizeof(GUID)))
		*ppvObject = &((_IOleClientSiteEx *)This)->inplace;

	else
	{
		*ppvObject = 0;
		return(E_NOINTERFACE);
	}

	return(S_OK);
}

HRESULT STDMETHODCALLTYPE SiteAddRef(IOleClientSite FAR* This)
{
	return(1);
}

HRESULT STDMETHODCALLTYPE SiteRelease(IOleClientSite FAR* This)
{
	return(1);
}

HRESULT STDMETHODCALLTYPE SiteSaveObject(IOleClientSite FAR* This)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE SiteGetMoniker(IOleClientSite FAR* This, DWORD dwAssign, DWORD dwWhichMoniker, IMoniker ** ppmk)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE SiteGetContainer(IOleClientSite FAR* This, LPOLECONTAINER FAR* ppContainer)
{
	*ppContainer = 0;

	return(E_NOINTERFACE);
}

HRESULT STDMETHODCALLTYPE SiteShowObject(IOleClientSite FAR* This)
{
	return(NOERROR);
}

HRESULT STDMETHODCALLTYPE SiteOnShowWindow(IOleClientSite FAR* This, BOOL fShow)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE SiteRequestNewObjectLayout(IOleClientSite FAR* This)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE SiteGetWindow(IOleInPlaceSite FAR* This, HWND FAR* lphwnd)
{
	*lphwnd = ((_IOleInPlaceSiteEx FAR*)This)->frame->window;

	return(S_OK);
}

HRESULT STDMETHODCALLTYPE SiteContextSensitiveHelp(IOleInPlaceSite FAR* This, BOOL fEnterMode)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE SiteCanInPlaceActivate(IOleInPlaceSite FAR* This)
{
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE SiteOnInPlaceActivate(IOleInPlaceSite FAR* This)
{
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE SiteOnUIActivate(IOleInPlaceSite FAR* This)
{
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE SiteGetWindowContext(IOleInPlaceSite FAR* This, LPOLEINPLACEFRAME FAR* lplpFrame, LPOLEINPLACEUIWINDOW FAR* lplpDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
	*lplpFrame = (LPOLEINPLACEFRAME)((_IOleInPlaceSiteEx FAR*)This)->frame;

	*lplpDoc = 0;
	lpFrameInfo->fMDIApp = FALSE;
	lpFrameInfo->hwndFrame = ((IOleInPlaceFrameEx FAR*)*lplpFrame)->window;
	lpFrameInfo->haccel = 0;
	lpFrameInfo->cAccelEntries = 0;

	GetClientRect(lpFrameInfo->hwndFrame, lprcPosRect);
	GetClientRect(lpFrameInfo->hwndFrame, lprcClipRect);
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE SiteScroll(IOleInPlaceSite FAR* This, SIZE scrollExtent)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE SiteOnUIDeactivate(IOleInPlaceSite FAR* This, BOOL fUndoable)
{
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE SiteOnInPlaceDeactivate(IOleInPlaceSite FAR* This)
{
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE SiteDiscardUndoState(IOleInPlaceSite FAR* This)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE SiteDeactivateAndUndo(IOleInPlaceSite FAR* This)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE SiteOnPosRectChange(IOleInPlaceSite FAR* This, LPCRECT lprcPosRect)
{
	return(S_OK);
}


HRESULT STDMETHODCALLTYPE FrameQueryInterface(IOleInPlaceFrame FAR* This, REFIID riid, LPVOID FAR* ppvObj)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE FrameAddRef(IOleInPlaceFrame FAR* This)
{
	return(1);
}

HRESULT STDMETHODCALLTYPE FrameRelease(IOleInPlaceFrame FAR* This)
{
	return(1);
}

HRESULT STDMETHODCALLTYPE FrameGetWindow(IOleInPlaceFrame FAR* This, HWND FAR* lphwnd)
{
	*lphwnd = ((IOleInPlaceFrameEx FAR*)This)->window;
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE FrameContextSensitiveHelp(IOleInPlaceFrame FAR* This, BOOL fEnterMode)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE FrameGetBorder(IOleInPlaceFrame FAR* This, LPRECT lprectBorder)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE FrameRequestBorderSpace(IOleInPlaceFrame FAR* This, LPCBORDERWIDTHS pborderwidths)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE FrameSetBorderSpace(IOleInPlaceFrame FAR* This, LPCBORDERWIDTHS pborderwidths)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE FrameSetActiveObject(IOleInPlaceFrame FAR* This, IOleInPlaceActiveObject *pActiveObject, LPCOLESTR pszObjName)
{
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE FrameInsertMenus(IOleInPlaceFrame FAR* This, HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE FrameSetMenu(IOleInPlaceFrame FAR* This, HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject)
{
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE FrameRemoveMenus(IOleInPlaceFrame FAR* This, HMENU hmenuShared)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE FrameSetStatusText(IOleInPlaceFrame FAR* This, LPCOLESTR pszStatusText)
{
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE FrameEnableModeless(IOleInPlaceFrame FAR* This, BOOL fEnable)
{
	return(S_OK);
}

HRESULT STDMETHODCALLTYPE FrameTranslateAccelerator(IOleInPlaceFrame FAR* This, LPMSG lpmsg, WORD wID)
{
	return(E_NOTIMPL);
}


void ReleaseBrowserObject(HWND hwnd)
{
	IOleObject	**browserHandle;
	IOleObject	*browserObject;

	if ((browserHandle = (IOleObject **)GetWindowLong(hwnd, GWL_USERDATA)))
	{
		browserObject = *browserHandle;
		browserObject->lpVtbl->Close(browserObject, OLECLOSE_NOSAVE);
		browserObject->lpVtbl->Release(browserObject);

		GlobalFree(browserHandle);

		return;
	}

}


void ConnectEventSink(IWebBrowser2* pSite,IEventHandler* pEventHandler)
{
	HRESULT hr;
	IConnectionPointContainer* pCPC;
	IConnectionPoint*          pCP;
	if (pSite)
	{
		hr = pSite->lpVtbl->QueryInterface(pSite, &IID_IConnectionPointContainer, (void**)&pCPC);
		if (FAILED(hr))
			return; 
		hr = pCPC->lpVtbl->FindConnectionPoint(pCPC, &DIID_DWebBrowserEvents2, &pCP);
		if (FAILED(hr)) { 
			pCPC->lpVtbl->Release(pCPC);
			return;
		}
		DWORD adviseCookie = 0;
		hr = pCP->lpVtbl->Advise(pCP, (IUnknown*)pEventHandler, &adviseCookie);
	}
}

long Navigate(HWND hwnd, LPCTSTR string, IEventHandler* pEventHandler)
{	
	IWebBrowser2	*webBrowser2;
	IOleObject		*browserObject;
	VARIANT			myURL;
	BSTR			bstr;

	browserObject = *((IOleObject **)GetWindowLong(hwnd, GWL_USERDATA));

	bstr = 0;

	if (!browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**)&webBrowser2))
	{
		VariantInit(&myURL);
		myURL.vt = VT_BSTR;
		myURL.bstrVal = SysAllocString(L"www.example.com");
		webBrowser2->lpVtbl->Navigate2(webBrowser2, &myURL, 0, 0, 0, 0);
		VariantClear(&myURL);
		ConnectEventSink(webBrowser2, pEventHandler);
		webBrowser2->lpVtbl->Release(webBrowser2);
	}
	return 0;
}


void DisplayHTMLPage(HWND hwnd, LPTSTR webPageName)
{
	IWebBrowser2	*webBrowser2;
	VARIANT			myURL;
	IOleObject		*browserObject;

	browserObject = *((IOleObject **)GetWindowLong(hwnd, GWL_USERDATA));

	if (!browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**)&webBrowser2))
	{
		VariantInit(&myURL);
		myURL.vt = VT_BSTR;

#ifndef UNICODE
		{
		wchar_t		*buffer;
		DWORD		size;

		size = MultiByteToWideChar(CP_ACP, 0, webPageName, -1, 0, 0);
		if (!(buffer = (wchar_t *)GlobalAlloc(GMEM_FIXED, sizeof(wchar_t) * size))) goto badalloc;
		MultiByteToWideChar(CP_ACP, 0, webPageName, -1, buffer, size);
		myURL.bstrVal = SysAllocString(buffer);
		GlobalFree(buffer);
		}
#else
		myURL.bstrVal = SysAllocString(webPageName);
#endif
		if (!myURL.bstrVal)
		{
badalloc:	webBrowser2->lpVtbl->Release(webBrowser2);
			return ;
		}

		webBrowser2->lpVtbl->Navigate2(webBrowser2, &myURL, 0, 0, 0, 0);
		VariantClear(&myURL);
		webBrowser2->lpVtbl->Release(webBrowser2);
	}

	return ;
}
void CreateBrowserObject(HWND hwnd)
{
	IOleObject			*browserObject;
	IWebBrowser2		*webBrowser2;
	RECT				rect;
	char				*ptr;
	IOleInPlaceFrameEx	*iOleInPlaceFrameEx;
	_IOleClientSiteEx	*_iOleClientSiteEx;

	if (!(ptr = (char *)GlobalAlloc(GMEM_FIXED, sizeof(IOleInPlaceFrameEx) + sizeof(_IOleClientSiteEx) + sizeof(IOleObject *))))
		return;
	
	iOleInPlaceFrameEx = (IOleInPlaceFrameEx *)(ptr + sizeof(IOleObject *));
	iOleInPlaceFrameEx->frame.lpVtbl = &MyIOleInPlaceFrameTable;

	iOleInPlaceFrameEx->window = hwnd;

	_iOleClientSiteEx = (_IOleClientSiteEx *)(ptr + sizeof(IOleInPlaceFrameEx) + sizeof(IOleObject *));
	_iOleClientSiteEx->client.lpVtbl = &MyIOleClientSiteTable;

	_iOleClientSiteEx->inplace.inplace.lpVtbl = &MyIOleInPlaceSiteTable;

	_iOleClientSiteEx->inplace.frame = iOleInPlaceFrameEx;

	if (!OleCreate(&CLSID_WebBrowser, &IID_IOleObject, OLERENDER_DRAW, 0, (IOleClientSite *)_iOleClientSiteEx, &MyIStorage, (void**)&browserObject))
	{
		*((IOleObject **)ptr) = browserObject;
		SetWindowLong(hwnd, GWL_USERDATA, (LONG)ptr);
		browserObject->lpVtbl->SetHostNames(browserObject, L"My Host Name", 0);

		GetClientRect(hwnd, &rect);

		if (!OleSetContainedObject((struct IUnknown *)browserObject, TRUE) &&
			!browserObject->lpVtbl->DoVerb(browserObject, OLEIVERB_SHOW, NULL, (IOleClientSite *)_iOleClientSiteEx, -1, hwnd, &rect) &&
			!browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**)&webBrowser2))
		{
			webBrowser2->lpVtbl->put_Left(webBrowser2, 0);
			webBrowser2->lpVtbl->put_Top(webBrowser2, 0);
			webBrowser2->lpVtbl->put_Width(webBrowser2, rect.right);
			webBrowser2->lpVtbl->put_Height(webBrowser2, rect.bottom);
			webBrowser2->lpVtbl->Release(webBrowser2);
			return;
		}
		ReleaseBrowserObject(hwnd);
		return;
	}

	GlobalFree(ptr);
	return;
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	if (uMsg == WM_CREATE)
	{
		CreateBrowserObject(hwnd);
		++wndCount;
		return(0);
	}

	if (uMsg == WM_DESTROY)
	{
		ReleaseBrowserObject(hwnd);
		--wndCount;
		if (!wndCount)
			PostQuitMessage(0);

		return(TRUE);
	}

	return(DefWindowProc(hwnd, uMsg, wParam, lParam));
}


int CALLBACK WinMain(HINSTANCE hInstance, HINSTANCE hInstNULL, LPSTR lpszCmdLine, int nCmdShow)
{
	MSG			msg;
	if (OleInitialize(NULL) == S_OK)
	{
		WNDCLASSEX		wc;
		ZeroMemory(&wc, sizeof(WNDCLASSEX));
		wc.cbSize = sizeof(WNDCLASSEX);
		wc.hInstance = hInstance;
		wc.lpfnWndProc = WindowProc;
		wc.lpszClassName = &className[0];
		RegisterClassEx(&wc);

		objectsCount = serverLockCount = 0;
		myFactory.lpVtbl = (IClassFactoryVtbl*)&IClassFactory_Vtbl;
		IClassFactory** pFactory = 0;
		HRESULT hr = FactoryQueryInterface(&myFactory, &IID_IEventHandler, pFactory);
		IEventHandler* eventHandler = 0;
		FactoryCreateInstance(&myFactory, 0, &IID_IEventHandler, &eventHandler);

		if ((msg.hwnd = CreateWindowEx(0, &className[0], "Html Page", WS_OVERLAPPEDWINDOW,
							CW_USEDEFAULT, 0, CW_USEDEFAULT, 0,
							HWND_DESKTOP, NULL, hInstance, 0)))
		{
			Navigate(msg.hwnd, "<H2>HTML Test</H2>", eventHandler);
			ShowWindow(msg.hwnd, nCmdShow);
			UpdateWindow(msg.hwnd);
		}

		while (GetMessage(&msg, 0, 0, 0))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		OleUninitialize();

		return 0;
	}

	return 0;
}



