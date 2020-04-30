#include "ClsFact.h"
#include "Guid.h"
#include "Globals.h"
#include "DeskBand.h"

CClassFactory::CClassFactory(CLSID clsid)
{
    m_clsidObject = clsid;
    m_ObjRefCount = 1;
    g_DllRefCount++;
}

CClassFactory::~CClassFactory()
{
    g_DllRefCount--;
}

STDMETHODIMP CClassFactory::QueryInterface(REFIID riid, LPVOID *ppReturn)
{
    *ppReturn = NULL;

    if (IsEqualIID(riid, IID_IUnknown))
        *ppReturn = this;
    else if (IsEqualIID(riid, IID_IClassFactory))
        *ppReturn = (IClassFactory*)this;

    if (*ppReturn)
    {
        (*(LPUNKNOWN*)ppReturn)->AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

STDMETHODIMP_(DWORD) CClassFactory::AddRef()
{
    return ++m_ObjRefCount;
}

STDMETHODIMP_(DWORD) CClassFactory::Release()
{
    if (--m_ObjRefCount == 0)
    {
        delete this;
        return 0;
    }

    return m_ObjRefCount;
}

STDMETHODIMP CClassFactory::CreateInstance(LPUNKNOWN pUnknown, REFIID riid, LPVOID *ppObject)
{
    HRESULT     hResult = E_FAIL;
    LPUNKNOWN   pTemp = NULL;

    *ppObject = NULL;

    if (pUnknown != NULL)
        return CLASS_E_NOAGGREGATION;

    if (IsEqualCLSID(m_clsidObject, CLSID_RadVDDBDeskBand))
    {
        CDeskBand *pDeskBand = new CDeskBand(CLSID_RadVDDBDeskBand);
        if (NULL == pDeskBand)
            return E_OUTOFMEMORY;

        pTemp = (LPUNKNOWN) (LPVOID) pDeskBand;
    }

    if (pTemp)
    {
        //get the QueryInterface return for our return value
        hResult = pTemp->QueryInterface(riid, ppObject);

        //call Release to decement the ref count
        pTemp->Release();
    }

    return hResult;
}

STDMETHODIMP CClassFactory::LockServer(BOOL)
{
    return E_NOTIMPL;
}
