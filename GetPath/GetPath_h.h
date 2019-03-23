

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 11:14:07 2038
 */
/* Compiler settings for GetPath.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.01.0622 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */


#ifndef __GetPath_h_h__
#define __GetPath_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IGetPath_FWD_DEFINED__
#define __IGetPath_FWD_DEFINED__
typedef interface IGetPath IGetPath;

#endif 	/* __IGetPath_FWD_DEFINED__ */


#ifndef __GetPath_FWD_DEFINED__
#define __GetPath_FWD_DEFINED__

#ifdef __cplusplus
typedef class GetPath GetPath;
#else
typedef struct GetPath GetPath;
#endif /* __cplusplus */

#endif 	/* __GetPath_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 



#ifndef __GetPath_LIBRARY_DEFINED__
#define __GetPath_LIBRARY_DEFINED__

/* library GetPath */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_GetPath;

#ifndef __IGetPath_DISPINTERFACE_DEFINED__
#define __IGetPath_DISPINTERFACE_DEFINED__

/* dispinterface IGetPath */
/* [uuid] */ 


EXTERN_C const IID DIID_IGetPath;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("30256D0E-AFEE-4389-9AF8-1AD85B590F6B")
    IGetPath : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IGetPathVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IGetPath * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IGetPath * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IGetPath * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IGetPath * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IGetPath * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IGetPath * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IGetPath * This,
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
        
        END_INTERFACE
    } IGetPathVtbl;

    interface IGetPath
    {
        CONST_VTBL struct IGetPathVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IGetPath_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IGetPath_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IGetPath_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IGetPath_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IGetPath_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IGetPath_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IGetPath_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IGetPath_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_GetPath;

#ifdef __cplusplus

class DECLSPEC_UUID("4F4D3B78-E2E9-45D2-987B-8C0FEEEDDDF9")
GetPath;
#endif
#endif /* __GetPath_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


