

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Thu Aug 06 21:32:39 2026
 */
/* Compiler settings for tsc64.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.00.0603 
    protocol : all , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
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


#ifndef __IScriptError_FWD_DEFINED__
#define __IScriptError_FWD_DEFINED__
typedef interface IScriptError IScriptError;

#endif 	/* __IScriptError_FWD_DEFINED__ */


#ifndef __IScriptControl_FWD_DEFINED__
#define __IScriptControl_FWD_DEFINED__
typedef interface IScriptControl IScriptControl;

#endif 	/* __IScriptControl_FWD_DEFINED__ */


#ifndef __IScriptError_FWD_DEFINED__
#define __IScriptError_FWD_DEFINED__
typedef interface IScriptError IScriptError;

#endif 	/* __IScriptError_FWD_DEFINED__ */


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
    {
        Initialized	= 0,
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
            _COM_Outptr_  void **ppvObject);
        
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
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
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


#ifndef __IScriptError_INTERFACE_DEFINED__
#define __IScriptError_INTERFACE_DEFINED__

/* interface IScriptError */
/* [oleautomation][dual][uuid][object] */ 


EXTERN_C const IID IID_IScriptError;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("70841C78-067D-11D0-95D8-00A02463AB28")
    IScriptError : public IDispatch
    {
    public:
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Number( 
            /* [retval][out] */ long *plNumber) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Source( 
            /* [retval][out] */ BSTR *pbstrSource) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Description( 
            /* [retval][out] */ BSTR *pbstrDescription) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_HelpFile( 
            /* [retval][out] */ BSTR *pbstrHelpFile) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_HelpContext( 
            /* [retval][out] */ long *plHelpContext) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Text( 
            /* [retval][out] */ BSTR *pbstrText) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Line( 
            /* [retval][out] */ long *plLine) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Column( 
            /* [retval][out] */ long *plColumn) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE Clear( void) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IScriptErrorVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IScriptError * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IScriptError * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IScriptError * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IScriptError * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IScriptError * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IScriptError * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IScriptError * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Number )( 
            IScriptError * This,
            /* [retval][out] */ long *plNumber);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Source )( 
            IScriptError * This,
            /* [retval][out] */ BSTR *pbstrSource);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Description )( 
            IScriptError * This,
            /* [retval][out] */ BSTR *pbstrDescription);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_HelpFile )( 
            IScriptError * This,
            /* [retval][out] */ BSTR *pbstrHelpFile);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_HelpContext )( 
            IScriptError * This,
            /* [retval][out] */ long *plHelpContext);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Text )( 
            IScriptError * This,
            /* [retval][out] */ BSTR *pbstrText);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Line )( 
            IScriptError * This,
            /* [retval][out] */ long *plLine);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Column )( 
            IScriptError * This,
            /* [retval][out] */ long *plColumn);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *Clear )( 
            IScriptError * This);
        
        END_INTERFACE
    } IScriptErrorVtbl;

    interface IScriptError
    {
        CONST_VTBL struct IScriptErrorVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IScriptError_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IScriptError_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IScriptError_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IScriptError_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IScriptError_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IScriptError_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IScriptError_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IScriptError_get_Number(This,plNumber)	\
    ( (This)->lpVtbl -> get_Number(This,plNumber) ) 

#define IScriptError_get_Source(This,pbstrSource)	\
    ( (This)->lpVtbl -> get_Source(This,pbstrSource) ) 

#define IScriptError_get_Description(This,pbstrDescription)	\
    ( (This)->lpVtbl -> get_Description(This,pbstrDescription) ) 

#define IScriptError_get_HelpFile(This,pbstrHelpFile)	\
    ( (This)->lpVtbl -> get_HelpFile(This,pbstrHelpFile) ) 

#define IScriptError_get_HelpContext(This,plHelpContext)	\
    ( (This)->lpVtbl -> get_HelpContext(This,plHelpContext) ) 

#define IScriptError_get_Text(This,pbstrText)	\
    ( (This)->lpVtbl -> get_Text(This,pbstrText) ) 

#define IScriptError_get_Line(This,plLine)	\
    ( (This)->lpVtbl -> get_Line(This,plLine) ) 

#define IScriptError_get_Column(This,plColumn)	\
    ( (This)->lpVtbl -> get_Column(This,plColumn) ) 

#define IScriptError_Clear(This)	\
    ( (This)->lpVtbl -> Clear(This) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IScriptError_INTERFACE_DEFINED__ */



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

unsigned long             __RPC_USER  BSTR_UserSize64(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal64(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal64(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree64(     unsigned long *, BSTR * ); 

unsigned long             __RPC_USER  LPSAFEARRAY_UserSize64(     unsigned long *, unsigned long            , LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserMarshal64(  unsigned long *, unsigned char *, LPSAFEARRAY * ); 
unsigned char * __RPC_USER  LPSAFEARRAY_UserUnmarshal64(unsigned long *, unsigned char *, LPSAFEARRAY * ); 
void                      __RPC_USER  LPSAFEARRAY_UserFree64(     unsigned long *, LPSAFEARRAY * ); 

unsigned long             __RPC_USER  VARIANT_UserSize64(     unsigned long *, unsigned long            , VARIANT * ); 
unsigned char * __RPC_USER  VARIANT_UserMarshal64(  unsigned long *, unsigned char *, VARIANT * ); 
unsigned char * __RPC_USER  VARIANT_UserUnmarshal64(unsigned long *, unsigned char *, VARIANT * ); 
void                      __RPC_USER  VARIANT_UserFree64(     unsigned long *, VARIANT * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


