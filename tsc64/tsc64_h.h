

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Tue Apr 28 21:49:20 2020
 */
/* Compiler settings for tsc64.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 7.00.0555 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __tsc64_h_h__
#define __tsc64_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IScriptControl_FWD_DEFINED__
#define __IScriptControl_FWD_DEFINED__
typedef interface IScriptControl IScriptControl;
#endif 	/* __IScriptControl_FWD_DEFINED__ */


#ifndef __IScriptControl_FWD_DEFINED__
#define __IScriptControl_FWD_DEFINED__
typedef interface IScriptControl IScriptControl;
#endif 	/* __IScriptControl_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IScriptControl_INTERFACE_DEFINED__
#define __IScriptControl_INTERFACE_DEFINED__

/* interface IScriptControl */
/* [dual][uuid][object] */ 

typedef /* [public][public][public] */ 
enum __MIDL_IScriptControl_0001
    {	Initialized	= 0,
	Connected	= 1
    } 	ScriptControlStates;


EXTERN_C const IID IID_IScriptControl;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0E59F1D3-1FBE-11D0-8FF2-00A0D10038BC")
    IScriptControl : public IDispatch
    {
    public:
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Language( 
            /* [retval][out] */ BSTR *pbstrLanguage) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Language( 
            /* [in] */ BSTR pbstrLanguage) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_State( 
            /* [retval][out] */ ScriptControlStates *pssState) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_State( 
            /* [in] */ ScriptControlStates pssState) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_SitehWnd( 
            /* [in] */ long phwnd) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_SitehWnd( 
            /* [retval][out] */ long *phwnd) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Timeout( 
            /* [retval][out] */ long *plMilleseconds) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_Timeout( 
            /* [in] */ long plMilleseconds) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_AllowUI( 
            /* [retval][out] */ VARIANT_BOOL *pfAllowUI) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_AllowUI( 
            /* [in] */ VARIANT_BOOL pfAllowUI) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_UseSafeSubset( 
            /* [retval][out] */ VARIANT_BOOL *pfUseSafeSubset) = 0;
        
        virtual /* [propput][id] */ HRESULT STDMETHODCALLTYPE put_UseSafeSubset( 
            /* [in] */ VARIANT_BOOL pfUseSafeSubset) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Modules( 
            /* [retval][out] */ IDispatch **ppmods) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Error( 
            /* [retval][out] */ IDispatch **ppse) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_CodeObject( 
            /* [retval][out] */ IDispatch **ppdispObject) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Procedures( 
            /* [retval][out] */ IDispatch **ppdispProcedures) = 0;
        
        virtual /* [hidden][id] */ HRESULT STDMETHODCALLTYPE _AboutBox( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE AddObject( 
            /* [in] */ BSTR Name,
            /* [in] */ IDispatch *Object,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL AddMembers = 0) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Reset( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE AddCode( 
            /* [in] */ BSTR Code) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Eval( 
            /* [in] */ BSTR Expression,
            /* [retval][out] */ VARIANT *pvarResult) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE ExecuteStatement( 
            /* [in] */ BSTR Statement) = 0;
        
        virtual /* [vararg][id] */ HRESULT STDMETHODCALLTYPE Run( 
            /* [in] */ BSTR ProcedureName,
            /* [in] */ SAFEARRAY * *Parameters,
            /* [retval][out] */ VARIANT *pvarResult) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IScriptControlVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IScriptControl * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IScriptControl * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IScriptControl * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IScriptControl * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IScriptControl * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IScriptControl * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IScriptControl * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Language )( 
            IScriptControl * This,
            /* [retval][out] */ BSTR *pbstrLanguage);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_Language )( 
            IScriptControl * This,
            /* [in] */ BSTR pbstrLanguage);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_State )( 
            IScriptControl * This,
            /* [retval][out] */ ScriptControlStates *pssState);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_State )( 
            IScriptControl * This,
            /* [in] */ ScriptControlStates pssState);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_SitehWnd )( 
            IScriptControl * This,
            /* [in] */ long phwnd);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_SitehWnd )( 
            IScriptControl * This,
            /* [retval][out] */ long *phwnd);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Timeout )( 
            IScriptControl * This,
            /* [retval][out] */ long *plMilleseconds);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_Timeout )( 
            IScriptControl * This,
            /* [in] */ long plMilleseconds);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_AllowUI )( 
            IScriptControl * This,
            /* [retval][out] */ VARIANT_BOOL *pfAllowUI);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_AllowUI )( 
            IScriptControl * This,
            /* [in] */ VARIANT_BOOL pfAllowUI);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_UseSafeSubset )( 
            IScriptControl * This,
            /* [retval][out] */ VARIANT_BOOL *pfUseSafeSubset);
        
        /* [propput][id] */ HRESULT ( STDMETHODCALLTYPE *put_UseSafeSubset )( 
            IScriptControl * This,
            /* [in] */ VARIANT_BOOL pfUseSafeSubset);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Modules )( 
            IScriptControl * This,
            /* [retval][out] */ IDispatch **ppmods);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Error )( 
            IScriptControl * This,
            /* [retval][out] */ IDispatch **ppse);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_CodeObject )( 
            IScriptControl * This,
            /* [retval][out] */ IDispatch **ppdispObject);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Procedures )( 
            IScriptControl * This,
            /* [retval][out] */ IDispatch **ppdispProcedures);
        
        /* [hidden][id] */ HRESULT ( STDMETHODCALLTYPE *_AboutBox )( 
            IScriptControl * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *AddObject )( 
            IScriptControl * This,
            /* [in] */ BSTR Name,
            /* [in] */ IDispatch *Object,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL AddMembers);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Reset )( 
            IScriptControl * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *AddCode )( 
            IScriptControl * This,
            /* [in] */ BSTR Code);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Eval )( 
            IScriptControl * This,
            /* [in] */ BSTR Expression,
            /* [retval][out] */ VARIANT *pvarResult);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *ExecuteStatement )( 
            IScriptControl * This,
            /* [in] */ BSTR Statement);
        
        /* [vararg][id] */ HRESULT ( STDMETHODCALLTYPE *Run )( 
            IScriptControl * This,
            /* [in] */ BSTR ProcedureName,
            /* [in] */ SAFEARRAY * *Parameters,
            /* [retval][out] */ VARIANT *pvarResult);
        
        END_INTERFACE
    } IScriptControlVtbl;

    interface IScriptControl
    {
        CONST_VTBL struct IScriptControlVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IScriptControl_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IScriptControl_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IScriptControl_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IScriptControl_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IScriptControl_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IScriptControl_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IScriptControl_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IScriptControl_get_Language(This,pbstrLanguage)	\
    ( (This)->lpVtbl -> get_Language(This,pbstrLanguage) ) 

#define IScriptControl_put_Language(This,pbstrLanguage)	\
    ( (This)->lpVtbl -> put_Language(This,pbstrLanguage) ) 

#define IScriptControl_get_State(This,pssState)	\
    ( (This)->lpVtbl -> get_State(This,pssState) ) 

#define IScriptControl_put_State(This,pssState)	\
    ( (This)->lpVtbl -> put_State(This,pssState) ) 

#define IScriptControl_put_SitehWnd(This,phwnd)	\
    ( (This)->lpVtbl -> put_SitehWnd(This,phwnd) ) 

#define IScriptControl_get_SitehWnd(This,phwnd)	\
    ( (This)->lpVtbl -> get_SitehWnd(This,phwnd) ) 

#define IScriptControl_get_Timeout(This,plMilleseconds)	\
    ( (This)->lpVtbl -> get_Timeout(This,plMilleseconds) ) 

#define IScriptControl_put_Timeout(This,plMilleseconds)	\
    ( (This)->lpVtbl -> put_Timeout(This,plMilleseconds) ) 

#define IScriptControl_get_AllowUI(This,pfAllowUI)	\
    ( (This)->lpVtbl -> get_AllowUI(This,pfAllowUI) ) 

#define IScriptControl_put_AllowUI(This,pfAllowUI)	\
    ( (This)->lpVtbl -> put_AllowUI(This,pfAllowUI) ) 

#define IScriptControl_get_UseSafeSubset(This,pfUseSafeSubset)	\
    ( (This)->lpVtbl -> get_UseSafeSubset(This,pfUseSafeSubset) ) 

#define IScriptControl_put_UseSafeSubset(This,pfUseSafeSubset)	\
    ( (This)->lpVtbl -> put_UseSafeSubset(This,pfUseSafeSubset) ) 

#define IScriptControl_get_Modules(This,ppmods)	\
    ( (This)->lpVtbl -> get_Modules(This,ppmods) ) 

#define IScriptControl_get_Error(This,ppse)	\
    ( (This)->lpVtbl -> get_Error(This,ppse) ) 

#define IScriptControl_get_CodeObject(This,ppdispObject)	\
    ( (This)->lpVtbl -> get_CodeObject(This,ppdispObject) ) 

#define IScriptControl_get_Procedures(This,ppdispProcedures)	\
    ( (This)->lpVtbl -> get_Procedures(This,ppdispProcedures) ) 

#define IScriptControl__AboutBox(This)	\
    ( (This)->lpVtbl -> _AboutBox(This) ) 

#define IScriptControl_AddObject(This,Name,Object,AddMembers)	\
    ( (This)->lpVtbl -> AddObject(This,Name,Object,AddMembers) ) 

#define IScriptControl_Reset(This)	\
    ( (This)->lpVtbl -> Reset(This) ) 

#define IScriptControl_AddCode(This,Code)	\
    ( (This)->lpVtbl -> AddCode(This,Code) ) 

#define IScriptControl_Eval(This,Expression,pvarResult)	\
    ( (This)->lpVtbl -> Eval(This,Expression,pvarResult) ) 

#define IScriptControl_ExecuteStatement(This,Statement)	\
    ( (This)->lpVtbl -> ExecuteStatement(This,Statement) ) 

#define IScriptControl_Run(This,ProcedureName,Parameters,pvarResult)	\
    ( (This)->lpVtbl -> Run(This,ProcedureName,Parameters,pvarResult) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IScriptControl_INTERFACE_DEFINED__ */



#ifndef __IScriptControl_LIBRARY_DEFINED__
#define __IScriptControl_LIBRARY_DEFINED__

/* library IScriptControl */
/* [uuid][version] */ 



EXTERN_C const IID LIBID_IScriptControl;
#endif /* __IScriptControl_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

unsigned long             __RPC_USER  LPSAFEARRAY_UserSize(     unsigned long *, unsigned long            , LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserMarshal(  unsigned long *, unsigned char *, LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserUnmarshal(unsigned long *, unsigned char *, LPSAFEARRAY * ); 
void                      __RPC_USER  LPSAFEARRAY_UserFree(     unsigned long *, LPSAFEARRAY * ); 

unsigned long             __RPC_USER  VARIANT_UserSize(     unsigned long *, unsigned long            , VARIANT * ); 
unsigned char * __RPC_USER  VARIANT_UserMarshal(  unsigned long *, unsigned char *, VARIANT * ); 
unsigned char * __RPC_USER  VARIANT_UserUnmarshal(unsigned long *, unsigned char *, VARIANT * ); 
void                      __RPC_USER  VARIANT_UserFree(     unsigned long *, VARIANT * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


