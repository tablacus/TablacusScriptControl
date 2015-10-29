// Tablacus Script Control 64 (C)2014 Gaku
// MIT Lisence
// Visual C++ 2008 Express Edition SP1
// Windows SDK v7.0
// http://www.eonet.ne.jp/~gakana/tablacus/

#include "tsc64.h"

#pragma comment (lib, "shlwapi.lib")
#ifdef _WIN64
const CLSID CLSID_TScriptServer = {0x0E59F1D5, 0x1FBE, 0x11D0, {0x8F, 0xF2, 0x00, 0xA0, 0xD1, 0x00, 0x38, 0xBC}};
const TCHAR g_szClsid[] = TEXT("{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}");
#else
const CLSID CLSID_TScriptServer = {0x0E59F1D5, 0x1FBE, 0x11D0, {0x8F, 0xF2, 0x00, 0xA0, 0xD1, 0x00, 0x38, 0xBD}};
const TCHAR g_szClsid[] = TEXT("{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BD}");
#endif

const CLSID IID_IScriptControl = {0x0E59F1D3, 0x1FBE, 0x11D0, {0x8F, 0xF2, 0x00, 0xA0, 0xD1, 0x00, 0x38, 0xBC}};
const CLSID DIID_DScriptControlSource = {0x8B167D60, 0x8605, 0x11D0, {0xAB, 0xCB, 0x00, 0xA0, 0xC9, 0x0F, 0xFF, 0xC0}};

const TCHAR g_szProgid[] = TEXT("Tablacus.ScriptControl");
LONG      g_lLocks = 0;
HINSTANCE g_hinstDll = NULL;
IScriptControl *pSC;


int GetIntFromVariant(VARIANT *pv)
{
	if (pv) {
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return GetIntFromVariant(pv->pvarVal);
		}
		if (pv->vt == VT_I4) {
			return pv->lVal;
		}
		if (pv->vt == VT_UI4) {
			return pv->ulVal;
		}
		if (pv->vt == VT_R8) {
			return (int)(LONGLONG)pv->dblVal;
		}
		VARIANT vo;
		VariantInit(&vo);
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I4)) {
			return vo.lVal;
		}
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_UI4)) {
			return vo.ulVal;
		}
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I8)) {
			return (int)vo.llVal;
		}
	}
	return 0;
}

int GetIntFromVariantClear(VARIANT *pv)
{
	int i = GetIntFromVariant(pv);
	VariantClear(pv);
	return i;
}

LONGLONG GetLLFromVariant(VARIANT *pv)
{
	if (pv) {
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return GetLLFromVariant(pv->pvarVal);
		}
		if (pv->vt == VT_I4) {
			return pv->lVal;
		}
		if (pv->vt == VT_R8) {
			return (LONGLONG)pv->dblVal;
		}
		VARIANT vo;
		VariantInit(&vo);
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I8)) {
			return vo.llVal;
		}
	}
	return 0;
}

VOID teSetLL(VARIANT *pv, LONGLONG ll)
{
	if (pv) {
		pv->lVal = static_cast<int>(ll);
		if (ll == static_cast<LONGLONG>(pv->lVal)) {
			pv->vt = VT_I4;
			return;
		}
		pv->dblVal = static_cast<DOUBLE>(ll);
		if (ll == static_cast<LONGLONG>(pv->dblVal)) {
			pv->vt = VT_R8;
			return;
		}
		pv->llVal = ll;
		pv->vt = VT_I8;
	}
}


VARIANTARG* GetNewVARIANT(int n)
{
	VARIANT *pv = new VARIANTARG[n];
	while (n--) {
		VariantInit(&pv[n]);
	}
	return pv;
}

BOOL FindUnknown(VARIANT *pv, IUnknown **ppunk)
{
	if (pv) {
		if (pv->vt == VT_DISPATCH || pv->vt == VT_UNKNOWN) {
			*ppunk = pv->punkVal;
			return *ppunk != NULL;
		}
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return FindUnknown(pv->pvarVal, ppunk);
		}
		if (pv->vt == (VT_DISPATCH | VT_BYREF) || pv->vt == (VT_UNKNOWN | VT_BYREF)) {
			*ppunk = *pv->ppunkVal;
			return *ppunk != NULL;
		}
	}
	*ppunk = NULL;
	return false;
}

HRESULT teCLSIDFromProgID(__in LPCOLESTR lpszProgID, __out LPCLSID lpclsid)
{
	if (StrCmpNI(lpszProgID, L"new:", 4) == 0) {
		return CLSIDFromString(&lpszProgID[4], lpclsid);
	}
	return CLSIDFromProgID(lpszProgID, lpclsid);
}

HRESULT Invoke5(IDispatch *pdisp, DISPID dispid, WORD wFlags, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs)
{
	HRESULT hr;

	int i;
	// DISPPARAMS
	DISPPARAMS dispParams;
	dispParams.rgvarg = pvArgs;
	dispParams.cArgs = abs(nArgs);
	DISPID dispidName = DISPID_PROPERTYPUT;
	if (wFlags & DISPATCH_PROPERTYPUT) {
		dispParams.cNamedArgs = 1;
		dispParams.rgdispidNamedArgs = &dispidName;
	} else {
		dispParams.rgdispidNamedArgs = NULL;
		dispParams.cNamedArgs = 0;
	}
	try {
		hr = pdisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
			wFlags, &dispParams, pvResult, NULL, NULL);
	} catch (...) {
		hr = E_FAIL;
	}
	// VARIANT Clean-up of an array
	if (pvArgs && nArgs >= 0) {
		for (i = nArgs - 1; i >=  0; i--){
			VariantClear(&pvArgs[i]);
		}
		delete[] pvArgs;
		pvArgs = NULL;
	}
	return hr;
}

HRESULT Invoke4(IDispatch *pdisp, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs)
{
	return Invoke5(pdisp, DISPID_VALUE, DISPATCH_METHOD, pvResult, nArgs, pvArgs);
}

BOOL teSetObject(VARIANT *pv, PVOID pObj)
{
	if (pObj) {
		try {
			IUnknown *punk = static_cast<IUnknown *>(pObj);
			if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->pdispVal))) {
				pv->vt = VT_DISPATCH;
				return true;
			}
			if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->punkVal))) {
				pv->vt = VT_UNKNOWN;
				return true;
			}
		} catch (...) {}
	}
	return false;
}

VOID teVariantChangeType(__out VARIANTARG * pvargDest,
				__in const VARIANTARG * pvarSrc, __in VARTYPE vt)
{
	VariantInit(pvargDest);
	if FAILED(VariantChangeType(pvargDest, pvarSrc, 0, vt)) {
		pvargDest->llVal = 0;
	}
}


// CTScriptControl

CTScriptControl::CTScriptControl()
{
	m_cRef = 1;
	Clear();
	m_pClientSite = NULL;
	LockModule(TRUE);
}

CTScriptControl::~CTScriptControl()
{
	raw_Reset();
	if (m_pClientSite) {
		m_pClientSite->Release();
	}
	LockModule(FALSE);
}

HRESULT CTScriptControl::Exec(BSTR Expression,VARIANT * pvarResult, DWORD dwFlags)
{
	if (m_pActiveScript) {
		IActiveScriptParse *pasp;
		if SUCCEEDED(m_pActiveScript->QueryInterface(IID_PPV_ARGS(&pasp))) {
			if (pasp->ParseScriptText(Expression, NULL, NULL, NULL, 0, 1, dwFlags, pvarResult, NULL) == S_OK) {
				m_pActiveScript->SetScriptState(SCRIPTSTATE_CONNECTED);
			}
			pasp->Release();
		}
	} else {
		ParseScript(Expression, m_bsLang, m_pObjectEx, NULL, &m_pCode, &m_pActiveScript, pvarResult, dwFlags);
	}
	return S_OK;
}

HRESULT CTScriptControl::ParseScript(LPOLESTR lpScript, LPOLESTR lpLang, IDispatchEx *pdex, IUnknown *pOnError, IDispatch **ppdisp, IActiveScript **ppas, VARIANT *pvarResult, DWORD dwFlags)
{
	HRESULT hr = E_FAIL;
	CLSID clsid;
	IActiveScript *pas = NULL;
	if (PathMatchSpec(lpLang, L"J*Script")) {
		CoCreateInstance(CLSID_JScriptChakra, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&pas));
	}
	if (pas == NULL && teCLSIDFromProgID(lpLang, &clsid) == S_OK) {
		CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&pas));
	}
	if (pas) {
		IActiveScriptProperty *paspr;
		if SUCCEEDED(pas->QueryInterface(IID_PPV_ARGS(&paspr))) {
			VARIANT v;
			VariantInit(&v);
			v.vt = VT_I4;
			v.lVal = 0;
			while (++v.lVal <= 256 && paspr->SetProperty(SCRIPTPROP_INVOKEVERSIONING, NULL, &v) == S_OK) {
			}
		}
		CteActiveScriptSite *pass = new CteActiveScriptSite(pdex, pOnError, this);
		pas->SetScriptSite(pass);
		pass->Release();
		IActiveScriptParse *pasp;
		if SUCCEEDED(pas->QueryInterface(IID_PPV_ARGS(&pasp))) {
			hr = pasp->InitNew();
			if (pdex) {
				DISPID dispid;
				HRESULT hr = pdex->GetNextDispID(fdexEnumAll, DISPID_STARTENUM, &dispid);
				while (hr == NOERROR) {
					BSTR bs;
					if (pdex->GetMemberName(dispid, &bs) == S_OK) {
						pas->AddNamedItem(bs, SCRIPTITEM_ISVISIBLE | SCRIPTITEM_ISSOURCE | SCRIPTITEM_GLOBALMEMBERS);
						::SysFreeString(bs);
					}
					hr = pdex->GetNextDispID(fdexEnumAll, dispid, &dispid);
				}
			}
			VARIANT v;
			VariantInit(&v);
			hr = pasp->ParseScriptText(lpScript, NULL, NULL, NULL, 0, 1, dwFlags, pvarResult, NULL);
			if (hr == S_OK) {
				pas->SetScriptState(SCRIPTSTATE_CONNECTED);
				if (ppdisp) {
					IDispatch *pdisp;
					if (pas->GetScriptDispatch(NULL, &pdisp) == S_OK) {
						CteDispatch *odisp = new CteDispatch(pdisp, 2);
						pdisp->Release();
						if SUCCEEDED(pas->QueryInterface(IID_PPV_ARGS(&odisp->m_pActiveScript))) {
							odisp->QueryInterface(IID_PPV_ARGS(ppdisp));
						}
						odisp->Release();
					}
				}
			}
			pasp->Release();
		}
		if (!ppdisp || !*ppdisp) {
			pas->SetScriptState(SCRIPTSTATE_CLOSED);
			pas->Close();
		}
		if (ppas) {
			*ppas = pas;
		} else {
			pas->Release();
		}
	}
	return hr;
}

STDMETHODIMP CTScriptControl::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	} else if (IsEqualIID(riid, IID_IScriptControl) || IsEqualIID(riid, DIID_DScriptControlSource) ||
		IsEqualIID(riid, CLSID_TScriptServer)) {
		*ppvObject = static_cast<IScriptControl *>(this);
	} else if (IsEqualIID(riid, IID_IOleObject)) {
		*ppvObject = static_cast<IOleObject *>(this);
	} else if (IsEqualIID(riid, IID_IOleControl)) {
		*ppvObject = static_cast<IOleControl *>(this);
	} else if (IsEqualIID(riid, IID_IPersistStreamInit)) {
		*ppvObject = static_cast<IPersistStreamInit *>(this);
	} else {
/*////
		TCHAR szKey[256];
		StringFromGUID2(riid, szKey, 40);
		MessageBox(0,szKey,0,0);
////*/
		return E_NOINTERFACE;
	}
	AddRef();
	
	return S_OK;
}

STDMETHODIMP_(ULONG) CTScriptControl::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CTScriptControl::Release()
{
	if (InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CTScriptControl::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CTScriptControl::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (lstrcmpi(*rgszNames, L"Language") == 0) {
		*rgDispId = 1500;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"State") == 0) {
		*rgDispId = 1501;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"SitehWnd") == 0) {
		*rgDispId = 1502;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"Timeout") == 0) {
		*rgDispId = 1503;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"AllowUI") == 0) {
		*rgDispId = 1504;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"UseSafeSubset") == 0) {
		*rgDispId = 1505;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"Modules") == 0) {
		*rgDispId = 1506;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"Error") == 0) {
		*rgDispId = 1507;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"CodeObject") == 0) {
		*rgDispId = 1000;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"Procedures") == 0) {
		*rgDispId = 1001;
		return S_OK;
	}
	//method
	if (lstrcmpi(*rgszNames, L"_AboutBox") == 0) {
		*rgDispId = -552;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"AddObject") == 0) {
		*rgDispId = 2500;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"Reset") == 0) {
		*rgDispId = 2501;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"AddCode") == 0) {
		*rgDispId = 2000;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"Eval") == 0) {
		*rgDispId = 2001;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"ExecuteStatement") == 0) {
		*rgDispId = 2002;
		return S_OK;
	}
	if (lstrcmpi(*rgszNames, L"Run") == 0) {
		*rgDispId = 2003;
		return S_OK;
	}
//	MessageBox(0, *rgszNames, 0, 0);
	return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP CTScriptControl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	VARIANT v;
	HRESULT hr = S_OK;

	switch (dispIdMember) {
		//Language
		case 1500:
			if (nArg >= 0) {
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				hr = put_Language(v.bstrVal);
				VariantClear(&v);
			}
			if (pVarResult) {
				hr = get_Language(&pVarResult->bstrVal);
				pVarResult->vt = VT_BSTR;
			}
			return hr;
		//State
		case 1501:
			if (nArg >= 0) {
				hr = put_State((ScriptControlStates)GetIntFromVariant(&pDispParams->rgvarg[nArg]));
			}
			if (pVarResult) {
				hr = get_State((ScriptControlStates *)&pVarResult->lVal);
				pVarResult->vt = VT_I4;
			}
			return hr;
		//SitehWnd
		case 1502:
			if (nArg >= 0) {
				m_hwnd.ll = GetLLFromVariant(&pDispParams->rgvarg[nArg]);
			}
			teSetLL(pVarResult, m_hwnd.ll);
			return hr;
		//Timeout
		case 1503:
			if (nArg >= 0) {
				hr = put_Timeout(GetIntFromVariant(&pDispParams->rgvarg[nArg]));
			}
			if (pVarResult) {
				hr = get_Timeout(&pVarResult->lVal);
				pVarResult->vt = VT_I4;
			}
			return hr;
		//AllowUI
		case 1504:
			return S_OK;
		//UseSafeSubset
		case 1505:
			return S_OK;
		//Modules
		case 1506:
			return E_NOTIMPL;
		//Error
		case 1507:
			return S_OK;
		//CodeObject
		case 1000:
			if (pVarResult && m_pCode) {
				if SUCCEEDED(m_pCode->QueryInterface(IID_PPV_ARGS(&pVarResult->pdispVal))) {
					pVarResult->vt = VT_DISPATCH;
				}
			}
			return S_OK;
		//Procedures
		case 1001:
			return E_NOTIMPL;
		//_AboutBox
		case -552:
			return raw__AboutBox();			
		//AddObject
		case 2500:
			if (nArg >= 1) {
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				IUnknown *punk;
				if (FindUnknown(&pDispParams->rgvarg[nArg - 1], &punk)) {
					IDispatch *pdisp;
					if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pdisp))) {
						hr = raw_AddObject(v.bstrVal, pdisp, nArg >= 2 && GetIntFromVariant(&pDispParams->rgvarg[nArg - 2]));
						pdisp->Release();
					}
				}
			}
			return hr;
		//Reset
		case 2501:
			return raw_Reset();
		//AddCode
		case 2000:
			if (nArg >= 0) {
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				hr = raw_AddCode(v.bstrVal);
				VariantClear(&v);
			}
			return hr;
		//Eval
		case 2001:
			if (nArg >= 0) {
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				hr = raw_Eval(v.bstrVal, pVarResult);
				VariantClear(&v);
			}
			return hr;
		//ExecuteStatement
		case 2002:
			if (nArg >= 0) {
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				hr = raw_ExecuteStatement(v.bstrVal);
				VariantClear(&v);
			}
			return hr;
		//Run
		case 2003:
			if (nArg >= 0) {
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				SAFEARRAY *psa = NULL;
				VARIANT *pv;
				if (nArg >= 1) {
					psa = ::SafeArrayCreateVector(VT_VARIANT, 0, nArg);
					if (::SafeArrayAccessData(psa, (LPVOID *)&pv) == S_OK) {
						for (int i = nArg; i--;) {
							::VariantCopy(&pv[i], &pDispParams->rgvarg[nArg - i - 1]);
						}
						::SafeArrayUnaccessData(psa);
					}
				}
				hr = raw_Run(v.bstrVal, &psa, pVarResult);
				if (psa) {
					::SafeArrayDestroy(psa);
				}
				VariantClear(&v);
			}
			return hr;
		//this
		case DISPID_VALUE:
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
	}//end_switch
/*	TCHAR szError[16];
	swprintf_s(szError, 16, TEXT("%x"), dispIdMember);
	MessageBox(NULL, (LPWSTR)szError, 0, 0);*/
	return DISP_E_MEMBERNOTFOUND;
}

STDMETHODIMP CTScriptControl::get_Language(BSTR * pbstrLanguage)
{
	*pbstrLanguage = ::SysAllocString(m_bsLang);
	return S_OK;
}

STDMETHODIMP CTScriptControl::put_Language(BSTR pbstrLanguage)
{
	if (m_bsLang) {
		::SysFreeString(m_bsLang);
	}
	m_bsLang = ::SysAllocString(pbstrLanguage);
	return S_OK;
}

STDMETHODIMP CTScriptControl::get_State(enum ScriptControlStates *pssState)
{
	return m_pActiveScript->GetScriptState((SCRIPTSTATE *)pssState);
}

STDMETHODIMP CTScriptControl::put_State(enum ScriptControlStates ssState)
{
	return m_pActiveScript->SetScriptState((SCRIPTSTATE)ssState);
}

STDMETHODIMP CTScriptControl::put_SitehWnd(long hwnd)
{
	m_hwnd.l[0] = hwnd;
	m_hwnd.l[1] = 0;
	return S_OK;
}

STDMETHODIMP CTScriptControl::get_SitehWnd(long *phwnd)
{
	*phwnd = m_hwnd.l[0];
	return m_hwnd.l[1] ? E_HANDLE : S_OK;
}

STDMETHODIMP CTScriptControl::get_Timeout(long *plMilleseconds)
{
	*plMilleseconds = 0;
	return S_OK;
}

STDMETHODIMP CTScriptControl::put_Timeout(long lMilleseconds)
{
	return S_OK;
}

STDMETHODIMP CTScriptControl::get_AllowUI(VARIANT_BOOL * pfAllowUI)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::put_AllowUI(VARIANT_BOOL pfAllowUI)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::get_UseSafeSubset(VARIANT_BOOL * pfUseSafeSubset)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::put_UseSafeSubset(VARIANT_BOOL pfUseSafeSubset)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::get_Modules(struct IScriptModuleCollection ** ppmods)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::get_Error(struct IScriptError ** ppse)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::get_CodeObject(IDispatch ** ppdispObject)
{
	return m_pCode->QueryInterface(IID_PPV_ARGS(ppdispObject));
}

STDMETHODIMP CTScriptControl::get_Procedures(struct IScriptProcedureCollection ** ppdispProcedures)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::raw__AboutBox()
{
	MessageBox(NULL, L"Tablacus Script Control 64 Version " _T(STRING(VER_Y)) L"." _T(STRING(VER_M)) L"." _T(STRING(VER_D)) L"." _T(STRING(VER_Z)), TITLE, MB_ICONINFORMATION | MB_OK);
	return S_OK;		
}

STDMETHODIMP CTScriptControl::raw_AddObject(BSTR Name, IDispatch * Object, VARIANT_BOOL AddMembers)
{
	VARIANT v;
	DISPID dispid;
	if (m_pObjectEx == NULL) {
		//Get JScript Object
		LPOLESTR lp = L"function o(){return {}}";
		if (ParseScript(lp, L"JScript", NULL, NULL, &m_pJS, NULL, NULL, SCRIPTTEXT_ISPERSISTENT | SCRIPTTEXT_ISVISIBLE) == S_OK) {
			lp = L"o";
			VariantInit(&v);
			if (m_pJS->GetIDsOfNames(IID_NULL, &lp, 1, LOCALE_USER_DEFAULT, &dispid) == S_OK) {
				if (Invoke5(m_pJS, dispid, DISPATCH_PROPERTYGET, &v, 0, NULL) == S_OK) {
					m_pObject = v.pdispVal;
					v.vt = VT_EMPTY;
					if (Invoke4(m_pObject, &v, 0, NULL) == S_OK) {
						v.pdispVal->QueryInterface(IID_PPV_ARGS(&m_pObjectEx));
					}
				}
			}
		}
		VariantClear(&v);
	}
	if (m_pObjectEx->GetDispID(Name, fdexNameEnsure, &dispid) == S_OK) {
		v.pdispVal = Object;
		v.vt = VT_DISPATCH;
		DISPID dispidName = DISPID_PROPERTYPUT;
		DISPPARAMS args;
		args.cArgs = 1;
		args.cNamedArgs = 1;
		args.rgdispidNamedArgs = &dispidName;
		args.rgvarg = &v;
		m_pObjectEx->InvokeEx(dispid, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUTREF, &args, NULL, NULL,NULL);
	}
	return S_OK;
}

VOID CTScriptControl::Clear()
{
	m_bsLang = NULL;
	m_pActiveScript = NULL;
	m_pJS = NULL;
	m_pObject = NULL;
	m_pObjectEx = NULL;
	m_pCode = NULL;
	m_hwnd.hwnd = NULL;
}

STDMETHODIMP CTScriptControl::raw_Reset()
{
	if (m_bsLang) {
		::SysFreeString(m_bsLang);
	}
	if (m_pCode) {
		m_pCode->Release();
	}
	if (m_pObjectEx) {
		m_pObjectEx->Release();
	}
	if (m_pObject) {
		m_pObject->Release();
	}
	if (m_pJS) {
		m_pJS->Release();
	}
	if (m_pActiveScript) {
		m_pActiveScript->Release();
	}
	Clear();
	return S_OK;
}

STDMETHODIMP CTScriptControl::raw_AddCode(BSTR Code)
{
	return Exec(Code, NULL, SCRIPTTEXT_ISPERSISTENT | SCRIPTTEXT_ISVISIBLE);
}

STDMETHODIMP CTScriptControl::raw_Eval(BSTR Expression, VARIANT *pvarResult)
{
	return Exec(Expression, pvarResult, SCRIPTTEXT_ISEXPRESSION);
}

STDMETHODIMP CTScriptControl::raw_ExecuteStatement(BSTR Statement)
{
	return Exec(Statement, NULL, SCRIPTTEXT_ISPERSISTENT);
}

STDMETHODIMP CTScriptControl::raw_Run(BSTR ProcedureName, SAFEARRAY ** Parameters, VARIANT *pvarResult)
{
	VARIANT *pv;
	VARIANT *pv2 = NULL;
	long nArg = 0;
	if (*Parameters) {
		::SafeArrayGetUBound(*Parameters, 1, &nArg);
		nArg++;
		pv2 = GetNewVARIANT(nArg);
		if (::SafeArrayAccessData(*Parameters, (LPVOID *)&pv) == S_OK) {
			for (int i = nArg; i--;) {
				pv2[nArg - i - 1] = pv[i];
			}
			::SafeArrayUnaccessData(*Parameters);
		}
	}
	DISPID dispid;
	if (m_pCode->GetIDsOfNames(IID_NULL, &ProcedureName, 1, LOCALE_USER_DEFAULT, &dispid) == S_OK) {
		Invoke5(m_pCode, dispid, DISPATCH_METHOD, pvarResult, -nArg, pv2);
	}
	if (pv2) {
		delete [] pv2;
	}
	return S_OK;
}

//IOleObject
STDMETHODIMP CTScriptControl::SetClientSite(IOleClientSite *pClientSite)
{
	if (pClientSite) {
		if (m_pClientSite) {
			m_pClientSite->Release();
		}
		pClientSite->QueryInterface(IID_PPV_ARGS(&m_pClientSite));
	}
	return S_OK;
}

STDMETHODIMP CTScriptControl::GetClientSite(IOleClientSite **ppClientSite)
{
	if (m_pClientSite) {
		return m_pClientSite->QueryInterface(IID_PPV_ARGS(ppClientSite));
	}
	*ppClientSite = NULL;
	return S_OK;
}

STDMETHODIMP CTScriptControl::SetHostNames(LPCOLESTR szContainerApp, LPCOLESTR szContainerObj)
{
	return S_OK;
}

STDMETHODIMP CTScriptControl::Close(DWORD dwSaveOption)
{
	return S_OK;
}

STDMETHODIMP CTScriptControl::SetMoniker(DWORD dwWhichMoniker, IMoniker *pmk)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::InitFromData(IDataObject *pDataObject, BOOL fCreation, DWORD dwReserved)
{
	return S_OK;
}

STDMETHODIMP CTScriptControl::GetClipboardData(DWORD dwReserved, IDataObject **ppDataObject)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::DoVerb(LONG iVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex, HWND hwndParent, LPCRECT lprcPosRect)
{
	return S_OK;
}

STDMETHODIMP CTScriptControl::EnumVerbs(IEnumOLEVERB **ppEnumOleVerb)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::Update(void)
{
	return S_OK;
}

STDMETHODIMP CTScriptControl::IsUpToDate(void)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::GetUserClassID(CLSID *pClsid)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::GetUserType(DWORD dwFormOfType, LPOLESTR *pszUserType)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::SetExtent(DWORD dwDrawAspect, SIZEL *psizel)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::GetExtent(DWORD dwDrawAspect, SIZEL *psizel)
{
	psizel->cx = 0;
	psizel->cy = 0;
	return S_OK;
}

STDMETHODIMP CTScriptControl::Advise(IAdviseSink *pAdvSink, DWORD *pdwConnection)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::Unadvise(DWORD dwConnection)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::EnumAdvise(IEnumSTATDATA **ppenumAdvise)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::GetMiscStatus(DWORD dwAspect, DWORD *pdwStatus)
{
	*pdwStatus = 0;
	return S_OK;
}

STDMETHODIMP CTScriptControl::SetColorScheme(LOGPALETTE *pLogpal)
{
	return E_NOTIMPL;
}

//IPersist
STDMETHODIMP CTScriptControl::GetClassID(CLSID *pClassID)
{
	return E_NOTIMPL;
}

//IOleControl
STDMETHODIMP CTScriptControl::GetControlInfo(CONTROLINFO *pCI)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::OnMnemonic(MSG *pMsg)
{
	return S_OK;
}

STDMETHODIMP CTScriptControl::OnAmbientPropertyChange(DISPID dispID)
{
	return S_OK;
}

STDMETHODIMP CTScriptControl::FreezeEvents(BOOL bFreeze)
{
	return E_NOTIMPL;
}

//IPersistStreamInit
STDMETHODIMP CTScriptControl::IsDirty(void)
{
	return S_OK;
}

STDMETHODIMP CTScriptControl::Load(LPSTREAM pStm)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::Save(LPSTREAM pStm, BOOL fClearDirty)
{
	return S_OK;
}

STDMETHODIMP CTScriptControl::GetSizeMax(ULARGE_INTEGER *pCbSize)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::InitNew(void)
{
	return S_OK;
}

// CTScriptControlFactory

STDMETHODIMP CTScriptControlFactory::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IClassFactory)) {
		*ppvObject = static_cast<IClassFactory *>(this);
	} else {
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
}

STDMETHODIMP_(ULONG) CTScriptControlFactory::AddRef()
{
	LockModule(TRUE);
	return 2;
}

STDMETHODIMP_(ULONG) CTScriptControlFactory::Release()
{
	LockModule(FALSE);
	return 1;
}

STDMETHODIMP CTScriptControlFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
{
	CTScriptControl *p;
	HRESULT   hr;

	*ppvObject = NULL;

	if (pUnkOuter != NULL) {
		return CLASS_E_NOAGGREGATION;
	}
	p = new CTScriptControl();
	if (p == NULL) {
		return E_OUTOFMEMORY;
	}
	hr = p->QueryInterface(riid, ppvObject);
	p->Release();

	return hr;
}

STDMETHODIMP CTScriptControlFactory::LockServer(BOOL fLock)
{
	LockModule(fLock);
	return S_OK;
}


// DLL Export


STDAPI DllCanUnloadNow(void)
{
	return g_lLocks == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	static CTScriptControlFactory serverFactory;
	HRESULT hr;

	*ppv = NULL;
	
	if (IsEqualCLSID(rclsid, CLSID_TScriptServer)) {
		hr = serverFactory.QueryInterface(riid, ppv);
	} else {
		hr = CLASS_E_CLASSNOTAVAILABLE;
	}
	return hr;
}

STDAPI DllRegisterServer(void)
{
	TCHAR szModulePath[MAX_PATH];
	TCHAR szKey[256];

	swprintf_s(szKey, 256, TEXT("CLSID\\%s"), g_szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, TEXT("Tablacus.ScriptControl"))) {
		return E_FAIL;
	}
	GetModuleFileName(g_hinstDll, szModulePath, sizeof(szModulePath) / sizeof(TCHAR));
	swprintf_s(szKey, 256, TEXT("CLSID\\%s\\InprocServer32"), g_szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, szModulePath)) {
		return E_FAIL;
	}
	swprintf_s(szKey, 256, TEXT("CLSID\\%s\\InprocServer32"), g_szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, TEXT("ThreadingModel"), TEXT("Apartment"))) {
		return E_FAIL;
	}
	swprintf_s(szKey, 256, TEXT("CLSID\\%s\\ProgID"), g_szClsid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, (LPTSTR)g_szProgid)) {
		return E_FAIL;
	}
	swprintf_s(szKey, 256, TEXT("%s"), g_szProgid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, TEXT("Tablacus Script Control 64"))) {
		return E_FAIL;
	}
	swprintf_s(szKey, 256, TEXT("%s\\CLSID"), g_szProgid);
	if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, (LPTSTR)g_szClsid)) {
		return E_FAIL;
	}
	return S_OK;
}

STDAPI DllUnregisterServer(void)
{
	TCHAR szKey[256];

	swprintf_s(szKey, TEXT("CLSID\\%s"), g_szClsid);
	SHDeleteKey(HKEY_CLASSES_ROOT, szKey);
	
	swprintf_s(szKey, TEXT("%s"), g_szProgid);
	SHDeleteKey(HKEY_CLASSES_ROOT, szKey);

	return S_OK;
}

BOOL WINAPI DllMain(HINSTANCE hinstDll, DWORD dwReason, LPVOID lpReserved)
{
	switch (dwReason) {

		case DLL_PROCESS_ATTACH:
			g_hinstDll = hinstDll;
			return TRUE;
		}
	return TRUE;
}

//CteDispatch

CteDispatch::CteDispatch(IDispatch *pDispatch, int nMode)
{
	m_cRef = 1;
	m_pDispatch = pDispatch;
	m_pDispatch->AddRef();
	m_nMode = nMode;
	m_dispIdMember = DISPID_UNKNOWN;
	m_pActiveScript = NULL;
	m_nIndex = 0;
}

CteDispatch::~CteDispatch()
{
	if (m_pDispatch) {
		m_pDispatch->Release();
	}
	if (m_pActiveScript) {
		m_pActiveScript->SetScriptState(SCRIPTSTATE_CLOSED);
		m_pActiveScript->Close();
		m_pActiveScript->Release();
	}
}

STDMETHODIMP CteDispatch::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch)) {
		*ppvObject = static_cast<IDispatch *>(this);
	} else if (IsEqualIID(riid, IID_IEnumVARIANT)) {
		*ppvObject = static_cast<IEnumVARIANT *>(this);
	} else if (m_nMode) {
		return m_pDispatch->QueryInterface(riid, ppvObject);
	} else {
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
}

STDMETHODIMP_(ULONG) CteDispatch::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteDispatch::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteDispatch::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteDispatch::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDispatch::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (m_nMode) {
		if (m_nMode == 2) {
			if (lstrcmp(*rgszNames, L"Item") == 0) {
				*rgDispId = DISPID_VALUE;
				return S_OK;
			}
		}
		return m_pDispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
	}
	return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP CteDispatch::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	if (m_nMode) {
		if (dispIdMember == DISPID_NEWENUM) {
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
		}
		if (dispIdMember == DISPID_VALUE) {
			int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
			if (nArg >= 0) {
				VARIANTARG *pv = GetNewVARIANT(1);
				teVariantChangeType(pv, &pDispParams->rgvarg[nArg], VT_BSTR);
				m_pDispatch->GetIDsOfNames(IID_NULL, &pv[0].bstrVal, 1, LOCALE_USER_DEFAULT, &dispIdMember);
				VariantClear(&pv[0]);
				if (nArg >= 1) {
					VariantCopy(&pv[0], &pDispParams->rgvarg[nArg - 1]);
					Invoke5(m_pDispatch, dispIdMember, DISPATCH_PROPERTYPUT, NULL, 1, pv);
				} else {
					delete [] pv;
				}
				if (pVarResult) {
					Invoke5(m_pDispatch, dispIdMember, DISPATCH_PROPERTYGET, pVarResult, 0, NULL);
				}
				return S_OK;
			}
			if (pVarResult) {
				teSetObject(pVarResult, this);
			}
			return S_OK;
		}
		return m_pDispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	}
	if (wFlags & DISPATCH_METHOD) {
		return m_pDispatch->Invoke(m_dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	}
	if (pVarResult) {
		teSetObject(pVarResult, this);
	}
	return S_OK;
}

STDMETHODIMP CteDispatch::Next(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched)
{
	if (rgVar) {
		if (m_nMode) {
			IDispatchEx *pdex = NULL;
			if SUCCEEDED(m_pDispatch->QueryInterface(IID_PPV_ARGS(&pdex))) {
				HRESULT hr;
				hr = pdex->GetNextDispID(0, m_dispIdMember, &m_dispIdMember);
				if (hr == S_OK) {
					hr = pdex->GetMemberName(m_dispIdMember, &rgVar->bstrVal);
					rgVar->vt = VT_BSTR;
				}
				pdex->Release();
				return hr;
			}
		}
		int nCount = 0;
		VARIANT v;
		if (Invoke5(m_pDispatch, DISPID_TE_COUNT, DISPATCH_PROPERTYGET, &v, 0, NULL) == S_OK) {
			nCount = GetIntFromVariantClear(&v);
		}
		if (m_nIndex < nCount) {
			if (pCeltFetched) {
				*pCeltFetched = 1;
			}
			VARIANTARG *pv = GetNewVARIANT(1);
			pv[0].vt = VT_I4;
			pv[0].lVal = m_nIndex++;
			return Invoke5(m_pDispatch, DISPID_TE_ITEM, DISPATCH_METHOD, rgVar, 1, pv);
		}	
	}
	return S_FALSE;
}

STDMETHODIMP CteDispatch::Skip(ULONG celt)
{
	m_nIndex += celt;
	return S_OK;
}

STDMETHODIMP CteDispatch::Reset(void)
{
	m_nIndex = 0;
	return S_OK;
}

STDMETHODIMP CteDispatch::Clone(IEnumVARIANT **ppEnum)
{
	if (ppEnum) {
		CteDispatch *pdisp = new CteDispatch(m_pDispatch, m_nMode);
		pdisp->m_dispIdMember = m_dispIdMember;
		pdisp->m_nMode = m_nMode;
		*ppEnum = static_cast<IEnumVARIANT *>(pdisp);
		return S_OK;
	}
	return E_POINTER;
}

//CteActiveScriptSite
CteActiveScriptSite::CteActiveScriptSite(IUnknown *punk, IUnknown *pOnError, CTScriptControl *pSC)
{
	m_cRef = 1;
	m_pDispatchEx = NULL;
	m_pOnError = NULL;
	m_pSC = NULL;
	if (punk) {
		punk->QueryInterface(IID_PPV_ARGS(&m_pDispatchEx));
	}
	if (pOnError) {
		pOnError->QueryInterface(IID_PPV_ARGS(&m_pOnError));
	}
	if (pSC) {
		pSC->AddRef();
		m_pSC = pSC;
	}
}

CteActiveScriptSite::~CteActiveScriptSite()
{
	m_pSC && m_pSC->Release();
	m_pOnError && m_pOnError->Release();
	m_pDispatchEx && m_pDispatchEx->Release();
}

STDMETHODIMP CteActiveScriptSite::QueryInterface(REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IActiveScriptSite)) {
		*ppvObject = static_cast<IActiveScriptSite *>(this);
	} else if (IsEqualIID(riid, IID_IActiveScriptSiteWindow)) {
		*ppvObject = static_cast<IActiveScriptSiteWindow *>(this);
	} else {
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
}

STDMETHODIMP_(ULONG) CteActiveScriptSite::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteActiveScriptSite::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteActiveScriptSite::GetLCID(LCID *plcid)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteActiveScriptSite::GetItemInfo(LPCOLESTR pstrName,
                   DWORD dwReturnMask,
                   IUnknown **ppiunkItem,
                   ITypeInfo **ppti)
{
	HRESULT hr = TYPE_E_ELEMENTNOTFOUND;
	if (dwReturnMask & SCRIPTINFO_IUNKNOWN) {
		if (m_pDispatchEx) {
			BSTR bs = ::SysAllocString(pstrName);
			DISPID dispid;
			if (m_pDispatchEx->GetDispID(bs, 0, &dispid) == S_OK) {
				DISPPARAMS noargs = { NULL, NULL, 0, 0 };
				VARIANT v;
				VariantInit(&v);
				if (m_pDispatchEx->InvokeEx(dispid, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noargs, &v, NULL, NULL) == S_OK) {
					if (FindUnknown(&v, ppiunkItem)) {
						hr = S_OK;
					} else {
						VariantClear(&v);
					}
				}
			}
			::SysFreeString(bs);
		}
	}
	return hr;
}

STDMETHODIMP CteActiveScriptSite::GetDocVersionString(BSTR *pbstrVersion)
{
	*pbstrVersion = SysAllocString(TITLE);
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::OnScriptError(IActiveScriptError *pscripterror)
{
	if (!pscripterror) {
		return E_POINTER;
	}
	EXCEPINFO ei;
	BSTR bs = NULL;
	DWORD dwSourceContext = 0;
    DWORD ulLineNumber = 0;
    LONG lCharacterPosition = 0;
	pscripterror->GetExceptionInfo(&ei);
	pscripterror->GetSourceLineText(&bs);
	pscripterror->GetSourcePosition(&dwSourceContext, &ulLineNumber, &lCharacterPosition);
	
	TCHAR szMessage[65536];

	swprintf_s(szMessage, 65536, TEXT("Line: %d\nCharacter: %d\nError: %s\nCode: %X\nSource: %s"), ulLineNumber, lCharacterPosition, ei.bstrDescription, ei.scode, ei.bstrSource);
	MessageBox(NULL, szMessage, TITLE, MB_OK | MB_ICONERROR);

	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::OnStateChange(SCRIPTSTATE ssScriptState)
{
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::OnScriptTerminate(const VARIANT *pvarResult,const EXCEPINFO *pexcepinfo)
{
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::OnEnterScript(void)
{
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::OnLeaveScript(void)
{
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::GetWindow(HWND *phwnd)
{
	*phwnd = m_pSC ? m_pSC->m_hwnd.hwnd : 0;
	return S_OK;
}

STDMETHODIMP CteActiveScriptSite::EnableModeless(BOOL fEnable)
{
	return E_NOTIMPL;
}


// Function

void LockModule(BOOL bLock)
{
	if (bLock) {
		InterlockedIncrement(&g_lLocks);
	} else {
		InterlockedDecrement(&g_lLocks);
	}
}


BOOL CreateRegistryKey(HKEY hKeyRoot, LPTSTR lpszKey, LPTSTR lpszValue, LPTSTR lpszData)
{
	HKEY  hKey;
	LONG  lResult;
	DWORD dwSize;

	lResult = RegCreateKeyEx(hKeyRoot, lpszKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
	if (lResult != ERROR_SUCCESS) {
/*		swprintf_s(lpszKey, 16, TEXT("%x"), lResult);
		MessageBox(NULL, (LPWSTR)lpszKey, (LPWSTR)g_szClsid, 0);*/
		return FALSE;
	}
	if (lpszData != NULL) {
		dwSize = (lstrlen(lpszData) + 1) * sizeof(TCHAR);
	} else {
		dwSize = 0;
	}
	RegSetValueEx(hKey, lpszValue, 0, REG_SZ, (LPBYTE)lpszData, dwSize);
	RegCloseKey(hKey);
	
	return TRUE;
}