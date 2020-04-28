#include "resource.h"
#include <windows.h>
#include <shlwapi.h>
#include <ActivScp.h>
#include <DispEx.h>
#include <tchar.h>
#import <msscript.ocx>
using namespace MSScriptControl;
#pragma comment (lib, "shlwapi.lib")

#define DISPID_TE_ITEM  0x3fffffff
#define DISPID_TE_COUNT 0x3ffffffe
#define DISPID_TE_INDEX 0x3ffffffd
#define DISPID_TE_MAX TE_VI
#define EVENT_COOKIE 100

union TUHWND {
	long l[2];
	HWND hwnd;
	LONGLONG ll;
};

struct TEmethod
{
	LONG   id;
	LPWSTR name;
};

class CteDispatch : public IDispatch, public IEnumVARIANT
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IEnumVARIANT
	STDMETHODIMP Next(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched);
	STDMETHODIMP Skip(ULONG celt);
	STDMETHODIMP Reset(void);
	STDMETHODIMP Clone(IEnumVARIANT **ppEnum);

	CteDispatch(IDispatch *pDispatch, int nMode);
	~CteDispatch();

	DISPID		m_dispIdMember;
	int			m_nIndex;
	IActiveScript *m_pActiveScript;
private:
	LONG		m_cRef;
	IDispatch	*m_pDispatch;
	int			m_nMode;//0: Clone 1:collection 2:ScriptDispatch
};

class CTScriptError : public IScriptError
{
public:
	EXCEPINFO m_EI;
	DWORD m_ulLine;
	LONG m_lColumn;
	BSTR m_bsText;
	//IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IScriptError
	STDMETHODIMP get_Number(long *plNumber);
	STDMETHODIMP get_Source(BSTR *pbstrSource);
	STDMETHODIMP get_Description (BSTR *pbstrDescription);
	STDMETHODIMP get_HelpFile(BSTR *pbstrHelpFile);
	STDMETHODIMP get_HelpContext(long *plHelpContext);
	STDMETHODIMP get_Text(BSTR *pbstrText);
	STDMETHODIMP get_Line(long *plLine);
	STDMETHODIMP get_Column(long *plColumn);
	STDMETHODIMP raw_Clear();

	CTScriptError();
	~CTScriptError();
private:
	LONG   m_cRef;
};

class CTScriptControl : public IScriptControl,
	public IOleObject, public IPersistStreamInit,
	public IOleControl,
	public IConnectionPointContainer, IConnectionPoint
{
public:
	TUHWND m_hwnd;
	CTScriptError *m_pError;
	EXCEPINFO *m_pEI;
	HRESULT m_hr;
	IDispatch *m_pEventSink;
	//IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IScriptContotol
	STDMETHODIMP get_Language(BSTR * pbstrLanguage);
	STDMETHODIMP put_Language(BSTR bstrLanguage);
	STDMETHODIMP get_State(enum ScriptControlStates * pssState);
	STDMETHODIMP put_State(enum ScriptControlStates ssState);
	STDMETHODIMP put_SitehWnd(long hwnd);
	STDMETHODIMP get_SitehWnd(long * phwnd);
	STDMETHODIMP get_Timeout(long * plMilleseconds);
	STDMETHODIMP put_Timeout(long lMilleseconds);
	STDMETHODIMP get_AllowUI(VARIANT_BOOL * pfAllowUI);
	STDMETHODIMP put_AllowUI(VARIANT_BOOL fAllowUI);
	STDMETHODIMP get_UseSafeSubset(VARIANT_BOOL * pfUseSafeSubset);
	STDMETHODIMP put_UseSafeSubset(VARIANT_BOOL pfUseSafeSubset);
	STDMETHODIMP get_Modules(struct IScriptModuleCollection ** ppmods);
	STDMETHODIMP get_Error(struct IScriptError ** ppse);
	STDMETHODIMP get_CodeObject(IDispatch ** ppdispObject);
	STDMETHODIMP get_Procedures(struct IScriptProcedureCollection ** ppdispProcedures);
	STDMETHODIMP raw__AboutBox();
	STDMETHODIMP raw_AddObject(BSTR Name,IDispatch * Object,VARIANT_BOOL AddMembers);
	STDMETHODIMP raw_Reset();
	STDMETHODIMP raw_AddCode(BSTR Code);
	STDMETHODIMP raw_Eval(BSTR Expression,VARIANT * pvarResult);
	STDMETHODIMP raw_ExecuteStatement(BSTR Statement);
	STDMETHODIMP raw_Run(BSTR ProcedureName,SAFEARRAY ** Parameters,VARIANT * pvarResult);
	//IOleObject
	STDMETHODIMP SetClientSite(IOleClientSite *pClientSite);
	STDMETHODIMP GetClientSite(IOleClientSite **ppClientSite);
	STDMETHODIMP SetHostNames(LPCOLESTR szContainerApp, LPCOLESTR szContainerObj);
	STDMETHODIMP Close(DWORD dwSaveOption);
	STDMETHODIMP SetMoniker(DWORD dwWhichMoniker, IMoniker *pmk);
	STDMETHODIMP GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk);
	STDMETHODIMP InitFromData(IDataObject *pDataObject, BOOL fCreation, DWORD dwReserved);
	STDMETHODIMP GetClipboardData(DWORD dwReserved, IDataObject **ppDataObject);
	STDMETHODIMP DoVerb(LONG iVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex, HWND hwndParent, LPCRECT lprcPosRect);
	STDMETHODIMP EnumVerbs(IEnumOLEVERB **ppEnumOleVerb);
	STDMETHODIMP Update(void);
	STDMETHODIMP IsUpToDate(void);
	STDMETHODIMP GetUserClassID(CLSID *pClsid);
	STDMETHODIMP GetUserType(DWORD dwFormOfType, LPOLESTR *pszUserType);
	STDMETHODIMP SetExtent(DWORD dwDrawAspect, SIZEL *psizel);
	STDMETHODIMP GetExtent(DWORD dwDrawAspect, SIZEL *psizel);
	STDMETHODIMP Advise(IAdviseSink *pAdvSink, DWORD *pdwConnection);
//	STDMETHODIMP Unadvise(DWORD dwConnection);
	STDMETHODIMP EnumAdvise(IEnumSTATDATA **ppenumAdvise);
	STDMETHODIMP GetMiscStatus(DWORD dwAspect, DWORD *pdwStatus);
	STDMETHODIMP SetColorScheme(LOGPALETTE *pLogpal);
	//IPersist
    STDMETHODIMP GetClassID(CLSID *pClassID);
	//IPersistStreamInit
	STDMETHODIMP IsDirty(void);
	STDMETHODIMP Load(LPSTREAM pStm);
	STDMETHODIMP Save(LPSTREAM pStm, BOOL fClearDirty);
	STDMETHODIMP GetSizeMax(ULARGE_INTEGER *pCbSize);
	STDMETHODIMP InitNew(void);
	//IOleControl
	STDMETHODIMP GetControlInfo(CONTROLINFO *pCI);
	STDMETHODIMP OnMnemonic(MSG *pMsg);
	STDMETHODIMP OnAmbientPropertyChange(DISPID dispID);
	STDMETHODIMP FreezeEvents(BOOL bFreeze);
	//IConnectionPointContainer
	STDMETHODIMP EnumConnectionPoints(IEnumConnectionPoints **ppEnum);
	STDMETHODIMP FindConnectionPoint(REFIID riid, IConnectionPoint **ppCP);
	//IConnectionPoint
	STDMETHODIMP GetConnectionInterface(IID *pIID);
	STDMETHODIMP GetConnectionPointContainer(IConnectionPointContainer **ppCPC);
	STDMETHODIMP Advise(IUnknown *pUnkSink, DWORD *pdwCookie);
	STDMETHODIMP Unadvise(DWORD dwCookie);
	STDMETHODIMP EnumConnections(IEnumConnections **ppEnum);

	CTScriptControl();
	~CTScriptControl();

	HRESULT Exec(BSTR Expression,VARIANT * pvarResult, DWORD dwFlags);
	HRESULT ParseScript(LPOLESTR lpScript, LPOLESTR lpLang, IDispatchEx *pdex, IDispatch **ppdisp, IActiveScript **ppas, VARIANT *pvarResult, DWORD dwFlags);
	VOID Clear();
	HRESULT SetScriptError(int n);

	VARIANT_BOOL m_fAllowUI;
	VARIANT_BOOL m_fUseSafeSubset;
private:
	LONG   m_cRef;
	BSTR m_bsLang;
	IActiveScript *m_pActiveScript;
	IDispatch *m_pJS;
	IDispatch *m_pObject;
	IDispatchEx *m_pObjectEx;
	IDispatch *m_pCode;
	IOleClientSite *m_pClientSite;
	ITypeInfo *m_pTypeInfo;
};

class CteActiveScriptSite : public IActiveScriptSite, public IActiveScriptSiteWindow
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//ActiveScriptSite
	STDMETHODIMP GetLCID(LCID *plcid);
	STDMETHODIMP GetItemInfo(LPCOLESTR pstrName,
                           DWORD dwReturnMask,
                           IUnknown **ppiunkItem,
                           ITypeInfo **ppti);
	STDMETHODIMP GetDocVersionString(BSTR *pbstrVersion);
	STDMETHODIMP OnScriptError(IActiveScriptError *pscripterror);
	STDMETHODIMP OnStateChange(SCRIPTSTATE ssScriptState);
	STDMETHODIMP OnScriptTerminate(const VARIANT *pvarResult,const EXCEPINFO *pexcepinfo);
	STDMETHODIMP OnEnterScript(void);
	STDMETHODIMP OnLeaveScript(void);
	//IActiveScriptSiteWindow
	STDMETHODIMP GetWindow(HWND *phwnd);
	STDMETHODIMP EnableModeless(BOOL fEnable);

	CteActiveScriptSite(IUnknown *punk, CTScriptControl *pSC);
	~CteActiveScriptSite();
public:
	IDispatchEx	*m_pDispatchEx;
	CTScriptControl *m_pSC;
	LONG		m_cRef;
};

class CTScriptControlFactory : public IClassFactory
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	STDMETHODIMP CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject);
	STDMETHODIMP LockServer(BOOL fLock);
};

class CTScriptObject : public IUnknown
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	CTScriptObject(IDispatch *pObject, VARIANT_BOOL bAddMembers);
	~CTScriptObject();

	IDispatch	*m_pObject;
	VARIANT_BOOL m_bAddMembers;
	LONG		m_cRef;
};

const CLSID CLSID_JScriptChakra       = {0x16d51579, 0xa30b, 0x4c8b, { 0xa2, 0x76, 0x0f, 0xf4, 0xdc, 0x41, 0xe7, 0x55}};
