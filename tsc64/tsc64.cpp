// Tablacus Script Control 64 (C)2014 Gaku
// MIT Lisence
// Visual C++ 2017 Express Edition
// Windows SDK v7.1
// https://tablacus.github.io/

#include "tsc64.h"

#ifdef _WIN64
const TCHAR g_szProgid[] = TEXT("MSScriptControl.ScriptControl");
const TCHAR g_szClsid[] = TEXT("{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}");
const TCHAR g_szLibid[] = TEXT("{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}");
const CLSID LIBID_TScriptControl = {0x0E59F1D2, 0x1FBE, 0x11D0, {0x8F, 0xF2, 0x00, 0xA0, 0xD1, 0x00, 0x38, 0xBC}};
#else
const TCHAR g_szProgid[] = TEXT("Tablacus.ScriptControl");
const TCHAR g_szClsid[] = TEXT("{760F48FE-E6E8-4d9d-AFD4-C7B393D4211F}");//test
const TCHAR g_szLibid[] = TEXT("{956BC468-C878-4BB4-BB0B-ACA410002E31}");//test
const CLSID LIBID_TScriptControl = {0x956BC468, 0xC878, 0x4BB4, {0xBB, 0x0B, 0xAC, 0xA4, 0x10, 0x00, 0x2E, 0x31}};//test
#endif

const CLSID IID_IScriptControl = {0x0E59F1D3, 0x1FBE, 0x11D0, {0x8F, 0xF2, 0x00, 0xA0, 0xD1, 0x00, 0x38, 0xBC}};
const CLSID DIID_DScriptControlSource = {0x8B167D60, 0x8605, 0x11D0, {0xAB, 0xCB, 0x00, 0xA0, 0xC9, 0x0F, 0xFF, 0xC0}};

LONG      g_lLocks = 0;
HINSTANCE g_hinstDll = NULL;
IScriptControl *pSC;
CLSID CLSID_TScriptServer;
GUID	g_ClsIdScriptObject;

int		*g_map;

TEmethod methodTSC[] = {
	//property
	{ 1500, L"Language" },
	{ 1501, L"State" },
	{ 1502, L"SitehWnd" },
	{ 1503, L"Timeout" },
	{ 1504, L"AllowUI" },
	{ 1505, L"UseSafeSubset" },
	{ 1506, L"Modules" },
	{ 1507, L"Error" },
	{ 1000, L"CodeObject" },
	{ 1001, L"Procedures" },
	//method
	{ -552, L"_AboutBox" },
	{ 2500, L"AddObject" },
	{ 2501, L"Reset" },
	{ 2000, L"AddCode" },
	{ 2001, L"Eval" },
	{ 2002, L"ExecuteStatement" },
	{ 2003, L"Run" },
	{ 0, NULL }
};

#pragma warning(push)
#pragma warning(disable:4838)
TEmethod methodTSE[] = {
	//property
	{ 0x000000c9, L"Number" },
	{ 0x000000ca, L"Source" },
	{ 0x000000cc, L"Description" },
	{ 0x000000cd, L"HelpContext" },
	{ 0xfffffdfb, L"Text" },
	{ 0x000000ce, L"Line" },
	{ 0xfffffdef, L"Column" },
	//method
	{ 0x000000d0, L"Clear" },
	{ 0, NULL }
};
#pragma warning(pop)

// Functions
VOID SafeRelease(PVOID ppObj)
{
	try {
		IUnknown **ppunk = static_cast<IUnknown **>(ppObj);
		if (*ppunk) {
			(*ppunk)->Release();
			*ppunk = NULL;
		}
	} catch (...) {}
}

VOID teSysFreeString(BSTR *pbs)
{
	if (*pbs) {
		::SysFreeString(*pbs);
		*pbs = NULL;
	}
}

void LockModule(BOOL bLock)
{
	if (bLock) {
		InterlockedIncrement(&g_lLocks);
	} else {
		InterlockedDecrement(&g_lLocks);
	}
}

VOID teClearExceptInfo(EXCEPINFO *pEI)
{
	teSysFreeString(&pEI->bstrSource);
	teSysFreeString(&pEI->bstrDescription);
	teSysFreeString(&pEI->bstrHelpFile);
	::ZeroMemory(pEI, sizeof(EXCEPINFO));
}

HRESULT ShowRegError(LSTATUS lr)
{
	LPTSTR lpBuffer = NULL;
	FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
		NULL, lr, LANG_USER_DEFAULT,
		(LPTSTR)&lpBuffer, 0, NULL );
	MessageBox(NULL, lpBuffer, g_szProgid, MB_ICONHAND | MB_OK);
	LocalFree(lpBuffer);
	return HRESULT_FROM_WIN32(lr);
}

LSTATUS CreateRegistryKey(HKEY hKeyRoot, LPTSTR lpszKey, LPTSTR lpszValue, LPTSTR lpszData)
{
	HKEY  hKey;
	LSTATUS  lr;
	DWORD dwSize;

	lr = RegCreateKeyEx(hKeyRoot, lpszKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
	if (lr == ERROR_SUCCESS) {
		if (lpszData != NULL) {
			dwSize = (lstrlen(lpszData) + 1) * sizeof(TCHAR);
		} else {
			dwSize = 0;
		}
		lr = RegSetValueEx(hKey, lpszValue, 0, REG_SZ, (LPBYTE)lpszData, dwSize);
		RegCloseKey(hKey);
	}
	return lr;
}

int* SortTEMethod(TEmethod *method, int nCount)
{
	int *pi = new int[nCount];

	for (int j = 0; j < nCount; j++) {
		BSTR bs = method[j].name;
		int nMin = 0;
		int nMax = j - 1;
		int nIndex;
		while (nMin <= nMax) {
			nIndex = (nMin + nMax) / 2;
			if (lstrcmpi(bs, method[pi[nIndex]].name) < 0) {
				nMax = nIndex - 1;
			} else {
				nMin = nIndex + 1;
			}
		}
		for (int i = j; i > nMin; i--) {
			pi[i] = pi[i - 1];
		}
		pi[nMin] = j;
	}
	return pi;
}

int teBSearch(TEmethod *method, int nSize, int* pMap, LPOLESTR bs)
{
	int nMin = 0;
	int nMax = nSize - 1;
	int nIndex, nCC;

	while (nMin <= nMax) {
		nIndex = (nMin + nMax) / 2;
		nCC = lstrcmpi(bs, method[pMap[nIndex]].name);
		if (nCC < 0) {
			nMax = nIndex - 1;
			continue;
		}
		if (nCC > 0) {
			nMin = nIndex + 1;
			continue;
		}
		return pMap[nIndex];
	}
	return -1;
}

HRESULT teGetDispId(TEmethod *method, int nCount, int* pMap, LPOLESTR bs, DISPID *rgDispId)
{
	if (pMap) {
		int nIndex = teBSearch(method, nCount, pMap, bs);
		if (nIndex >= 0) {
			*rgDispId = method[nIndex].id;
			return S_OK;
		}
	} else {
		for (int i = 0; method[i].name; i++) {
			if (lstrcmpi(bs, method[i].name) == 0) {
				*rgDispId = method[i].id;
				return S_OK;
			}
		}
	}
	return DISP_E_UNKNOWNNAME;
}

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

VOID teSetLong(VARIANT *pv, LONG i)
{
	if (pv) {
		pv->lVal = i;
		pv->vt = VT_I4;
	}
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

static void threadTimer(void *args)
{
	::OleInitialize(NULL);
	TETimer *pCommon = (TETimer *)args;
	try {
		BOOL bAbort = FALSE;
		if (WaitForSingleObject(pCommon->hEvent, pCommon->lTimeout) == WAIT_TIMEOUT) {
			bAbort = MessageBox(NULL, L"Timeout!\nDo you want to stop?", TITLE, MB_YESNO) == IDYES;
		}
		if (::InterlockedDecrement(&pCommon->cRef)) {
			if (bAbort) {
				IActiveScript *pActiveScript;
				if SUCCEEDED(CoGetInterfaceAndReleaseStream(pCommon->pStream, IID_PPV_ARGS(&pActiveScript))) {
					EXCEPINFO ei;
					HRESULT hr = pActiveScript->InterruptScriptThread(SCRIPTTHREADID_ALL, &ei, SCRIPTINTERRUPT_RAISEEXCEPTION);
					pActiveScript->Release();
					TCHAR szDebug[256];
					wsprintf(szDebug, TEXT("Tablacus Script Abort: %x\n"), hr);
					::OutputDebugString(szDebug);
				}
			} else {
				pCommon->pStream->Release();
			}
		} else {
			CloseHandle(pCommon->hEvent);
			pCommon->pStream->Release();
			delete pCommon;
		}
	} catch (...) {
	}
	::OleUninitialize();
	::_endthread();
}

// CTScriptControl

CTScriptControl::CTScriptControl()
{
	ITypeLib *pTypeLib;

	m_cRef = 1;
	Clear();
	m_bsLang = NULL;
	m_pClientSite = NULL;
	m_pTypeInfo = NULL;
	m_pEventSink = NULL;
	if SUCCEEDED(LoadRegTypeLib(LIBID_TScriptControl, 1, 0, 0, &pTypeLib)) {
		pTypeLib->GetTypeInfoOfGuid(IID_IScriptControl, &m_pTypeInfo);
		pTypeLib->Release();
	}
	m_pError = new CTScriptError();
	LockModule(TRUE);
}

CTScriptControl::~CTScriptControl()
{
	m_pEI = NULL;
	teSysFreeString(&m_bsLang);
	raw_Reset();
	SafeRelease(&m_pClientSite);
	SafeRelease(&m_pTypeInfo);
	SafeRelease(&m_pError);
	SafeRelease(&m_pEventSink);
	LockModule(FALSE);
}

HRESULT CTScriptControl::Exec(BSTR Expression,VARIANT * pvarResult, DWORD dwFlags)
{
	m_hr = S_OK;
	if (m_pActiveScript) {
		TETimer *pCommon = NULL;
		if (m_lTimeout) {
			IStream *pStream;
			if SUCCEEDED(CoMarshalInterThreadInterfaceInStream(IID_IActiveScript, m_pActiveScript, &pStream)) {
				pCommon = new TETimer[1];
				pCommon->pStream = pStream;
				pCommon->cRef = 2;
				pCommon->hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
				pCommon->lTimeout = m_lTimeout;
				_beginthread(&threadTimer, 0, pCommon);
			}
		}
		IActiveScriptParse *pasp;
		if SUCCEEDED(m_pActiveScript->QueryInterface(IID_PPV_ARGS(&pasp))) {
			if (pasp->ParseScriptText(Expression, NULL, NULL, NULL, 0, 1, dwFlags, pvarResult, NULL) == S_OK) {
				m_pActiveScript->SetScriptState(SCRIPTSTATE_CONNECTED);
			}
			pasp->Release();
		}
		if (pCommon) {
			if (::InterlockedDecrement(&pCommon->cRef)) {
				SetEvent(pCommon->hEvent);
			} else {
				CloseHandle(pCommon->hEvent);
				delete pCommon;
			}
		}
	} else {
		ParseScript(Expression, m_bsLang, m_pObjectEx, &m_pCode, &m_pActiveScript, pvarResult, dwFlags);
	}
	return m_hr;
}

HRESULT CTScriptControl::ParseScript(LPOLESTR lpScript, LPOLESTR lpLang, IDispatchEx *pdex, IDispatch **ppdisp, IActiveScript **ppas, VARIANT *pvarResult, DWORD dwFlags)
{
	HRESULT hr = E_FAIL;
	CLSID clsid;
	IActiveScript *pas = NULL;
	if (PathMatchSpec(lpLang, L"J*Script9")) {
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
		CteActiveScriptSite *pass = new CteActiveScriptSite(pdex, this);
		pas->SetScriptSite(pass);
		pass->Release();
		TETimer *pCommon = NULL;
		if (m_lTimeout) {
			IStream *pStream;
			if SUCCEEDED(CoMarshalInterThreadInterfaceInStream(IID_IActiveScript, pas, &pStream)) {
				pCommon = new TETimer[1];
				pCommon->pStream = pStream;
				pCommon->cRef = 2;
				pCommon->hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
				pCommon->lTimeout = m_lTimeout;
				_beginthread(&threadTimer, 0, pCommon);
			}
		}
		IActiveScriptParse *pasp;
		if SUCCEEDED(pas->QueryInterface(IID_PPV_ARGS(&pasp))) {
			hr = pasp->InitNew();
			if (pdex) {
				DISPID dispid;
				HRESULT hr = pdex->GetNextDispID(fdexEnumAll, DISPID_STARTENUM, &dispid);
				while (hr == NOERROR) {
					BSTR bs;
					if (pdex->GetMemberName(dispid, &bs) == S_OK) {
						DWORD dwFlags = SCRIPTITEM_ISPERSISTENT | SCRIPTITEM_ISVISIBLE;
						DISPPARAMS noargs = { NULL, NULL, 0, 0 };
						VARIANT v;
						VariantInit(&v);
						if (pdex->InvokeEx(dispid, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noargs, &v, NULL, NULL) == S_OK) {
							IUnknown *punk;
							if (FindUnknown(&v, &punk)) {
								CTScriptObject *pSO;
								if SUCCEEDED(punk->QueryInterface(g_ClsIdScriptObject, (LPVOID *)&pSO)) {
									if (pSO->m_bAddMembers) {
										dwFlags = SCRIPTITEM_ISVISIBLE | SCRIPTITEM_ISSOURCE | SCRIPTITEM_GLOBALMEMBERS;
									}
									pSO->Release();
								}
							}
						}
						VariantClear(&v);
						pas->AddNamedItem(bs, dwFlags);
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
		if (pCommon) {
			if (::InterlockedDecrement(&pCommon->cRef)) {
				SetEvent(pCommon->hEvent);
			} else {
				CloseHandle(pCommon->hEvent);
				delete pCommon;
			}
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
#pragma warning(push)
#pragma warning(disable:4838)
	static const QITAB qit[] =
	{
		QITABENT(CTScriptControl, IDispatch),
		QITABENT(CTScriptControl, IScriptControl),
		&DIID_DScriptControlSource, OFFSETOFCLASS(IScriptControl, CTScriptControl),
		&CLSID_TScriptServer, OFFSETOFCLASS(IScriptControl, CTScriptControl),
		QITABENT(CTScriptControl, IOleObject),
		QITABENT(CTScriptControl, IOleControl),
		QITABENT(CTScriptControl, IPersistStreamInit),
		QITABENT(CTScriptControl, IConnectionPointContainer),
		QITABENT(CTScriptControl, IConnectionPoint),
		{ 0 },
	};
#pragma warning(pop)
	return QISearch(this, qit, riid, ppvObject);
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
	*pctinfo = m_pTypeInfo ? 1 : 0;
	return S_OK;
}

STDMETHODIMP CTScriptControl::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	if (m_pTypeInfo) {
		m_pTypeInfo->AddRef();
		*ppTInfo = m_pTypeInfo;
		return S_OK;
	}
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	return teGetDispId(methodTSC, _countof(methodTSC), g_map, *rgszNames, rgDispId);
}

STDMETHODIMP CTScriptControl::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	VARIANT v; VariantInit(&v);
	HRESULT hr = S_OK;

	// used by CteActiveScriptSite::OnScriptError to pass info to our caller
	m_pEI = pExcepInfo;

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
			break;
		//State
		case 1501:
			if (nArg >= 0) {
				hr = put_State((ScriptControlStates)GetIntFromVariant(&pDispParams->rgvarg[nArg]));
			}
			if (pVarResult) {
				hr = get_State((ScriptControlStates *)&pVarResult->lVal);
				pVarResult->vt = VT_I4;
			}
			break;
		//SitehWnd
		case 1502:
			if (nArg >= 0) {
				m_hwnd.ll = GetLLFromVariant(&pDispParams->rgvarg[nArg]);
			}
			teSetLL(pVarResult, m_hwnd.ll);
			break;
		//Timeout
		case 1503:
			if (nArg >= 0) {
				hr = put_Timeout(GetIntFromVariant(&pDispParams->rgvarg[nArg]));
			}
			if (pVarResult) {
				hr = get_Timeout(&pVarResult->lVal);
				pVarResult->vt = VT_I4;
			}
			break;
		//AllowUI
		case 1504:
			if (nArg >= 0) {
				hr = put_AllowUI(GetIntFromVariant(&pDispParams->rgvarg[nArg]));
			}
			if (pVarResult) {
				hr = get_AllowUI(&pVarResult->boolVal);
				pVarResult->vt = VT_BOOL;
			}
			break;
		//UseSafeSubset
		case 1505:
			if (nArg >= 0) {
				hr = put_UseSafeSubset(GetIntFromVariant(&pDispParams->rgvarg[nArg]));
			}
			if (pVarResult) {
				hr = get_UseSafeSubset(&pVarResult->boolVal);
				pVarResult->vt = VT_BOOL;
			}
			break;
		//Modules
		case 1506:
			hr = E_NOTIMPL;
			break;
		//Error
		case 1507:
			teSetObject(pVarResult, m_pError);
			hr = S_OK;
			break;
		//CodeObject
		case 1000:
			if (pVarResult && m_pCode) {
				if SUCCEEDED(m_pCode->QueryInterface(IID_PPV_ARGS(&pVarResult->pdispVal))) {
					pVarResult->vt = VT_DISPATCH;
				}
			}
			hr = S_OK;
			break;
		//Procedures
		case 1001:
			hr = E_NOTIMPL;
			break;
		//_AboutBox
		case -552:
			hr = raw__AboutBox();
			break;
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
			break;
		//Reset
		case 2501:
			hr = raw_Reset();
			break;
		//AddCode
		case 2000:
			if (nArg >= 0) {
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				hr = raw_AddCode(v.bstrVal);
				VariantClear(&v);
			}
			break;
		//Eval
		case 2001:
			if (nArg >= 0) {
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				hr = raw_Eval(v.bstrVal, pVarResult);
				VariantClear(&v);
			}
			break;
		//ExecuteStatement
		case 2002:
			if (nArg >= 0) {
				teVariantChangeType(&v, &pDispParams->rgvarg[nArg], VT_BSTR);
				hr = raw_ExecuteStatement(v.bstrVal);
				VariantClear(&v);
			}
			break;
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
			break;
		//DIID_DScriptControlSource
/*		case 3000://Error
		case 3001://Timeout
			if (m_pEventSink) {
				hr = m_pEventSink->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
			} else {
				hr = DISP_E_MEMBERNOTFOUND;
			}
			break;*/
			
		case DISPID_VALUE://this
			teSetObject(pVarResult, this);
			hr = S_OK;
			break;
		default:
			hr = DISP_E_MEMBERNOTFOUND;
			break;
	}//end_switch
	
	// no more caller for CteActiveScriptSite::OnScriptError
	m_pEI = NULL;
	return hr;
}

STDMETHODIMP CTScriptControl::get_Language(BSTR * pbstrLanguage)
{
	*pbstrLanguage = ::SysAllocString(m_bsLang);
	return S_OK;
}

STDMETHODIMP CTScriptControl::put_Language(BSTR pbstrLanguage)
{
	teSysFreeString(&m_bsLang);
	m_bsLang = ::SysAllocString(pbstrLanguage);
	return S_OK;
}

STDMETHODIMP CTScriptControl::get_State(enum ScriptControlStates *pssState)
{
	return m_pActiveScript ? m_pActiveScript->GetScriptState((SCRIPTSTATE *)pssState) : E_FAIL;
}

STDMETHODIMP CTScriptControl::put_State(enum ScriptControlStates ssState)
{
	return m_pActiveScript ? m_pActiveScript->SetScriptState((SCRIPTSTATE)ssState) : E_FAIL;
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
	*plMilleseconds = m_lTimeout;
	return S_OK;
}

STDMETHODIMP CTScriptControl::put_Timeout(long lMilleseconds)
{
	m_lTimeout = lMilleseconds;
	return S_OK;
}

STDMETHODIMP CTScriptControl::get_AllowUI(VARIANT_BOOL * pfAllowUI)
{
	*pfAllowUI = m_fAllowUI;
	return S_OK;
}

STDMETHODIMP CTScriptControl::put_AllowUI(VARIANT_BOOL  fAllowUI)
{
	m_fAllowUI = fAllowUI;
	return S_OK;
}

STDMETHODIMP CTScriptControl::get_UseSafeSubset(VARIANT_BOOL * pfUseSafeSubset)
{
	*pfUseSafeSubset = m_fUseSafeSubset;
	return S_OK;
}

STDMETHODIMP CTScriptControl::put_UseSafeSubset(VARIANT_BOOL pfUseSafeSubset)
{
	m_fUseSafeSubset = pfUseSafeSubset;
	return S_OK;
}

STDMETHODIMP CTScriptControl::get_Modules(struct IScriptModuleCollection ** ppmods)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::get_Error(struct IScriptError ** ppse)
{
	*ppse = m_pError;
	return S_OK;
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
	MessageBox(NULL, _T(PRODUCTNAME) L" Version " _T(STRING(VER_Y)) L"." _T(STRING(VER_M)) L"." _T(STRING(VER_D)) L"." _T(STRING(VER_Z)), TITLE, MB_ICONINFORMATION | MB_OK);
	return S_OK;
}

STDMETHODIMP CTScriptControl::raw_AddObject(BSTR Name, IDispatch * Object, VARIANT_BOOL AddMembers)
{
	VARIANT v;
	DISPID dispid;
	if (m_pObjectEx == NULL) {
		//Get JScript Object
		LPOLESTR lp = L"function o(){return {}}";
		if (ParseScript(lp, L"JScript", NULL, &m_pJS, NULL, NULL, SCRIPTTEXT_ISPERSISTENT | SCRIPTTEXT_ISVISIBLE) == S_OK) {
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
		v.punkVal = new CTScriptObject(Object, AddMembers);
		v.vt = VT_UNKNOWN;
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
	m_pActiveScript = NULL;
	m_pJS = NULL;
	m_pObject = NULL;
	m_pObjectEx = NULL;
	m_pCode = NULL;
	m_hwnd.hwnd = NULL;
	m_fAllowUI = VARIANT_TRUE;
	m_fUseSafeSubset = VARIANT_FALSE;
	m_lTimeout = 0;
}

HRESULT CTScriptControl::SetScriptError(int n)
{
	HRESULT hr = E_UNEXPECTED;
	WCHAR pszBuff[4096];
	if (m_pEI) {
		lstrcpy(pszBuff, L"%SystemRoot%\\SysWOW64\\msscript.ocx");
		ExpandEnvironmentStrings(pszBuff, pszBuff, 4096);
		HMODULE hLib = LoadLibraryEx(pszBuff, NULL, LOAD_LIBRARY_AS_DATAFILE);
		if (hLib) {
			if (LoadString(hLib, n, pszBuff, 4096)) {
				m_pEI->bstrDescription = ::SysAllocString(pszBuff);
				m_pEI->bstrSource = ::SysAllocString(TITLE);
				m_pEI->scode = E_FAIL;
			}
			FreeLibrary(hLib);
		}
		hr = DISP_E_EXCEPTION;
	}
	return hr;
}

STDMETHODIMP CTScriptControl::raw_Reset()
{
	SafeRelease(&m_pCode);
	SafeRelease(&m_pObjectEx);
	SafeRelease(&m_pObject);
	SafeRelease(&m_pJS);
	SafeRelease(&m_pActiveScript);
	SafeRelease(&m_pEventSink);
	Clear();
	return m_bsLang ? S_OK : SetScriptError(1024);
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
	m_hr = S_OK;
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
	HRESULT hr = m_pCode->GetIDsOfNames(IID_NULL, &ProcedureName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (hr == S_OK) {
		Invoke5(m_pCode, dispid, DISPATCH_METHOD, pvarResult, -nArg, pv2);
	}
	if (pv2) {
		delete [] pv2;
	}
	return hr ? hr : m_hr;
}

//IOleObject
STDMETHODIMP CTScriptControl::SetClientSite(IOleClientSite *pClientSite)
{
	if (pClientSite) {
		SafeRelease(&m_pClientSite);
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
/* 
STDMETHODIMP CTScriptControl::Unadvise(DWORD dwConnection)
{
	return E_NOTIMPL;
}
*/
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

//IConnectionPointContainer
STDMETHODIMP CTScriptControl::EnumConnectionPoints(IEnumConnectionPoints **ppEnum)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::FindConnectionPoint(REFIID riid, IConnectionPoint **ppCP)
{
	return QueryInterface(IID_PPV_ARGS(ppCP));
}

//IConnectionPoint
STDMETHODIMP CTScriptControl::GetConnectionInterface(IID *pIID)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptControl::GetConnectionPointContainer(IConnectionPointContainer **ppCPC)
{
	return QueryInterface(IID_PPV_ARGS(ppCPC));
}

STDMETHODIMP CTScriptControl::Advise(IUnknown *pUnkSink, DWORD *pdwCookie)
{
	if (!pdwCookie) {
		return E_POINTER;
	}
	SafeRelease(&m_pEventSink);
	pUnkSink->QueryInterface(IID_PPV_ARGS(&m_pEventSink));
	if (m_pEventSink) {
		*pdwCookie = EVENT_COOKIE;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP CTScriptControl::Unadvise(DWORD dwCookie)
{
	if (dwCookie == EVENT_COOKIE) {
		SafeRelease(&m_pEventSink);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP CTScriptControl::EnumConnections(IEnumConnections **ppEnum)
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
#pragma warning(push)
#pragma warning(disable:4838)
	static const QITAB qit[] =
	{
		QITABENT(CTScriptControlFactory, IClassFactory),
		{ 0 },
	};
#pragma warning(pop)
	return QISearch(this, qit, riid, ppvObject);
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

//CTScriptObject
STDMETHODIMP CTScriptObject::QueryInterface(REFIID riid, void **ppvObject)
{
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, g_ClsIdScriptObject)) {
		*ppvObject = this;
		AddRef();
		return S_OK;
	}
	return E_NOINTERFACE;
}


STDMETHODIMP_(ULONG) CTScriptObject::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CTScriptObject::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}
	return m_cRef;
}

CTScriptObject::CTScriptObject(IDispatch *pObject, VARIANT_BOOL bAddMembers)
{
	m_cRef = 1;
	pObject->QueryInterface(IID_PPV_ARGS(&m_pObject));
	m_bAddMembers = bAddMembers;
}

CTScriptObject::~CTScriptObject()
{
	SafeRelease(&m_pObject);
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

	wsprintf(szKey, TEXT("CLSID\\%s"), g_szClsid);
	LSTATUS lr = CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, TEXT("ScriptControl Object"));
	if (lr != ERROR_SUCCESS) {
		return ShowRegError(lr);
	}
	GetModuleFileName(g_hinstDll, szModulePath, sizeof(szModulePath) / sizeof(TCHAR));
	wsprintf(szKey, TEXT("CLSID\\%s\\InprocServer32"), g_szClsid);
	lr = CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, szModulePath);
	if (lr != ERROR_SUCCESS) {
		return ShowRegError(lr);
	}
	lr = CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, TEXT("ThreadingModel"), TEXT("Apartment"));
	if (lr != ERROR_SUCCESS) {
		return ShowRegError(lr);
	}
	wsprintf(szKey, TEXT("CLSID\\%s\\ProgID"), g_szClsid);
	lr = CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, const_cast<LPTSTR>(g_szProgid));
	if (lr != ERROR_SUCCESS) {
		return ShowRegError(lr);
	}
	lr = CreateRegistryKey(HKEY_CLASSES_ROOT, const_cast<LPTSTR>(g_szProgid), NULL, TEXT(PRODUCTNAME));
	if (lr != ERROR_SUCCESS) {
		return ShowRegError(lr);
	}
	wsprintf(szKey, TEXT("%s\\CLSID"), g_szProgid);
	lr = CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, const_cast<LPTSTR>(g_szClsid));
	if (lr != ERROR_SUCCESS) {
		return ShowRegError(lr);
	}
	ITypeLib *pTypeLib;
	HRESULT hr = LoadTypeLib(szModulePath, &pTypeLib);
	if FAILED(hr) {
		ShowRegError(hr);
		return hr;
	}
	hr = RegisterTypeLib(pTypeLib, szModulePath, NULL);
	pTypeLib->Release();
	if FAILED(hr) {
		ShowRegError(hr);
		return hr;
	}
	wsprintf(szKey, TEXT("CLSID\\%s\\TypeLib"), g_szClsid);
	lr = CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, const_cast<LPTSTR>(g_szLibid));
	if (lr != ERROR_SUCCESS) {
		return ShowRegError(lr);
	}
	wsprintf(szKey, TEXT("CLSID\\%s\\Version"), g_szClsid);
	lr = CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, NULL, TEXT("1.0"));
	if (lr != ERROR_SUCCESS) {
		return ShowRegError(lr);
	}
	return S_OK;
}

STDAPI DllUnregisterServer(void)
{
	TCHAR szKey[64];
	wsprintf(szKey, TEXT("CLSID\\%s"), g_szClsid);
	LSTATUS ls = SHDeleteKey(HKEY_CLASSES_ROOT, szKey);
#ifdef _WIN64
	if (ls == ERROR_SUCCESS) {
		return S_OK;
	}
#else
	if (ls == ERROR_SUCCESS) {
		ls = SHDeleteKey(HKEY_CLASSES_ROOT, g_szProgid);
		if (ls == ERROR_SUCCESS) {
			return S_OK;
		}
	}
#endif
	return ShowRegError(ls);
}

BOOL WINAPI DllMain(HINSTANCE hinstDll, DWORD dwReason, LPVOID lpReserved)
{
	switch (dwReason) {

		case DLL_PROCESS_ATTACH:
			g_hinstDll = hinstDll;
			CLSIDFromString(g_szClsid, &CLSID_TScriptServer);
			g_map = SortTEMethod(methodTSC, _countof(methodTSC));
			CoCreateGuid(&g_ClsIdScriptObject);
			break;
		case DLL_PROCESS_DETACH:
			delete [] g_map;
			break;
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
	SafeRelease(&m_pDispatch);
	if (m_pActiveScript) {
		m_pActiveScript->SetScriptState(SCRIPTSTATE_CLOSED);
		m_pActiveScript->Close();
		m_pActiveScript->Release();
	}
}

STDMETHODIMP CteDispatch::QueryInterface(REFIID riid, void **ppvObject)
{
#pragma warning(push)
#pragma warning(disable:4838)
	static const QITAB qit[] =
	{
		QITABENT(CteDispatch, IDispatch),
		QITABENT(CteDispatch, IEnumVARIANT),
		{ 0 },
	};
#pragma warning(pop)
	HRESULT hr = QISearch(this, qit, riid, ppvObject);
	if SUCCEEDED(hr) {
		return hr;
	}
	if (m_nMode) {
		return m_pDispatch->QueryInterface(riid, ppvObject);
	}
	return E_NOINTERFACE;
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
			teSetObject(pVarResult, this);
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
			teSetObject(pVarResult, this);
			return S_OK;
		}
		return m_pDispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	}
	if (wFlags & DISPATCH_METHOD) {
		return m_pDispatch->Invoke(m_dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	}
	teSetObject(pVarResult, this);
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
CteActiveScriptSite::CteActiveScriptSite(IUnknown *punk, CTScriptControl *pSC)
{
	m_cRef = 1;
	m_pDispatchEx = NULL;
	m_pSC = NULL;
	if (punk) {
		punk->QueryInterface(IID_PPV_ARGS(&m_pDispatchEx));
	}
	if (pSC) {
		pSC->AddRef();
		m_pSC = pSC;
	}
}

CteActiveScriptSite::~CteActiveScriptSite()
{
	SafeRelease(&m_pSC);
	SafeRelease(&m_pDispatchEx);
}

STDMETHODIMP CteActiveScriptSite::QueryInterface(REFIID riid, void **ppvObject)
{
#pragma warning(push)
#pragma warning(disable:4838)
	static const QITAB qit[] =
	{
		QITABENT(CteActiveScriptSite, IActiveScriptSite),
		QITABENT(CteActiveScriptSite, IActiveScriptSiteWindow),
		{ 0 },
	};
#pragma warning(pop)
	return QISearch(this, qit, riid, ppvObject);
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
						CTScriptObject *pSO;
						if SUCCEEDED((*ppiunkItem)->QueryInterface(g_ClsIdScriptObject, (LPVOID *)&pSO)) {
							SafeRelease(ppiunkItem);
							*ppiunkItem = pSO->m_pObject;
						}
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
	EXCEPINFO *pei = &m_pSC->m_pError->m_EI;
	teClearExceptInfo(pei);
	if SUCCEEDED(pscripterror->GetExceptionInfo(pei)) {
		DWORD dwSourceContext = 0;
		pscripterror->GetSourcePosition(&dwSourceContext, &m_pSC->m_pError->m_ulLine, &m_pSC->m_pError->m_lColumn);
		teSysFreeString(&m_pSC->m_pError->m_bsText);
		pscripterror->GetSourceLineText(&m_pSC->m_pError->m_bsText);
		if (m_pSC->m_pEventSink) {
			VARIANT v;
			VariantInit(&v);
			if (Invoke5(m_pSC->m_pEventSink, 3000, DISPATCH_METHOD, &v, 0, NULL) == S_OK) {
				VariantClear(&v);
			}
		}
		if (m_pSC->m_pEI) {
			if SUCCEEDED(pscripterror->GetExceptionInfo(m_pSC->m_pEI)) {
				m_pSC->m_hr = DISP_E_EXCEPTION;
			}
		}
	}
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

//CTScriptError

CTScriptError::CTScriptError()
{
	m_cRef = 1;
	m_EI.bstrSource = NULL;
	m_EI.bstrDescription = NULL;
	m_EI.bstrHelpFile = NULL;
	m_bsText = NULL;
	raw_Clear();
}

CTScriptError::~CTScriptError()
{
	raw_Clear();
}

STDMETHODIMP CTScriptError::QueryInterface(REFIID riid, void **ppvObject)
{
#pragma warning(push)
#pragma warning(disable:4838)
	static const QITAB qit[] =
	{
		QITABENT(CTScriptError, IDispatch),
		QITABENT(CTScriptError, IScriptError),
		{ 0 },
	};
#pragma warning(pop)
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CTScriptError::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CTScriptError::Release()
{
	if (InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CTScriptError::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CTScriptError::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CTScriptError::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	return teGetDispId(methodTSE, _countof(methodTSE), NULL, *rgszNames, rgDispId);
}

STDMETHODIMP CTScriptError::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	switch (dispIdMember) {
		//Number
		case 0x000000c9:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				return get_Number(&pVarResult->lVal);
			}
			return S_OK;
		//Source
		case 0x000000ca:
			if (pVarResult) {
				pVarResult->vt = VT_BSTR;
				return get_Source(&pVarResult->bstrVal);
			}
			return S_OK;
		//Description
		case 0x000000cc:
			if (pVarResult) {
				pVarResult->vt = VT_BSTR;
				return get_Description(&pVarResult->bstrVal);
			}
			return S_OK;
		//HelpContext
		case 0x000000cd:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				return get_HelpContext(&pVarResult->lVal);
			}
			return S_OK;
		//Text
		case 0xfffffdfb:
			if (pVarResult) {
				pVarResult->vt = VT_BSTR;
				return get_Text(&pVarResult->bstrVal);
			}
			return S_OK;
		//Line
		case 0x000000ce:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				return get_Line(&pVarResult->lVal);
			}
			return S_OK;
		//Column
		case 0xfffffdef:
			if (pVarResult) {
				pVarResult->vt = VT_I4;
				return get_Column(&pVarResult->lVal);
			}
			return S_OK;
		//Clear
		case 0x000000d0:
			return raw_Clear();
		//this
		case DISPID_VALUE:
			teSetObject(pVarResult, this);
			return S_OK;
	}//end_switch
	return DISP_E_MEMBERNOTFOUND;
}

STDMETHODIMP CTScriptError::get_Number(long *plNumber)
{
	*plNumber = LOWORD(m_EI.scode);
	return S_OK;
}

STDMETHODIMP CTScriptError::get_Source(BSTR *pbstrSource)
{
	*pbstrSource = ::SysAllocString(m_EI.bstrSource);
	return S_OK;
}

STDMETHODIMP CTScriptError::get_Description (BSTR *pbstrDescription)
{
	*pbstrDescription = ::SysAllocString(m_EI.bstrDescription);
	return S_OK;
}

STDMETHODIMP CTScriptError::get_HelpFile(BSTR *pbstrHelpFile)
{
	*pbstrHelpFile = ::SysAllocString(m_EI.bstrHelpFile);
	return S_OK;
}

STDMETHODIMP CTScriptError::get_HelpContext(long *plHelpContext)
{
	*plHelpContext = m_EI.dwHelpContext;
	return S_OK;
}

STDMETHODIMP CTScriptError::get_Text(BSTR *pbstrText)
{
	*pbstrText = ::SysAllocString(m_bsText);
	return S_OK;
}

STDMETHODIMP CTScriptError::get_Line(long *plLine)
{
	*plLine = m_ulLine;
	return S_OK;
}

STDMETHODIMP CTScriptError::get_Column(long *plColumn)
{
	*plColumn = m_lColumn;
	return S_OK;
}

STDMETHODIMP CTScriptError::raw_Clear()
{
	m_ulLine = 0;
	m_lColumn = 0;
	teSysFreeString(&m_bsText);
	teClearExceptInfo(&m_EI);
	return S_OK;
}
