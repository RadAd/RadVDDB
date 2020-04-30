#include <ole2.h>
#include <comcat.h>
#include <olectl.h>
#include <tchar.h>
#include "ClsFact.h"

//this part is only done once
//if you need to use the GUID in another file, just include Guid.h
#pragma data_seg(".text")
#define INITGUID
#include <initguid.h>
#include <shlguid.h>
#include "Guid.h"
#pragma data_seg()

extern "C" BOOL WINAPI DllMain(HINSTANCE, DWORD, LPVOID);
BOOL RegisterServer(CLSID, LPTSTR);
BOOL UnregisterServer(CLSID);
BOOL RegisterComCat(CLSID, CATID);
BOOL UnregisterComCat(CLSID, CATID);


HINSTANCE   g_hInst;
UINT        g_DllRefCount;


extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    switch(dwReason)
    {
    case DLL_PROCESS_ATTACH:
        g_hInst = hInstance;
        break;

    case DLL_PROCESS_DETACH:
        break;
    }

    return TRUE;
}

STDAPI DllCanUnloadNow(void)
{
    return (g_DllRefCount ? S_FALSE : S_OK);
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppReturn)
{
    *ppReturn = NULL;

    // if we don't support this classid, return the proper error code
    if (!IsEqualCLSID(rclsid, CLSID_RadVDDBDeskBand))
        return CLASS_E_CLASSNOTAVAILABLE;

    //create a CClassFactory object and check it for validity
    CClassFactory *pClassFactory = new CClassFactory(rclsid);
    if (NULL == pClassFactory)
        return E_OUTOFMEMORY;

    //get the QueryInterface return for our return value
    HRESULT hResult = pClassFactory->QueryInterface(riid, ppReturn);

    //call Release to decement the ref count - creating the object set it to one
    //and QueryInterface incremented it - since its being used externally (not by
    //us), we only want the ref count to be 1
    pClassFactory->Release();

    //return the result from QueryInterface
    return hResult;
}

STDAPI DllRegisterServer(void)
{
    if (!RegisterServer(CLSID_RadVDDBDeskBand, TEXT("Rad &Virtual Desktop")))
        return SELFREG_E_CLASS;
    if (!RegisterComCat(CLSID_RadVDDBDeskBand, CATID_DeskBand))
        return SELFREG_E_CLASS;

    return S_OK;
}

STDAPI DllUnregisterServer(void)
{
    if (!UnregisterComCat(CLSID_RadVDDBDeskBand, CATID_DeskBand))
        return SELFREG_E_CLASS;
    if (!UnregisterServer(CLSID_RadVDDBDeskBand))
        return SELFREG_E_CLASS;

    return S_OK;
}

typedef struct{
    HKEY    hRootKey;
    LPTSTR    szSubKey;//TCHAR szSubKey[MAX_PATH];
    LPTSTR    lpszValueName;
    LPTSTR    szData;//TCHAR szData[MAX_PATH];
}    DOREGSTRUCT, *LPDOREGSTRUCT;

BOOL RegisterServer(CLSID clsid, LPTSTR lpszTitle)
{
    int      i;
    HKEY     hKey;
    LRESULT  lResult;
    DWORD    dwDisp;
    TCHAR    szSubKey[MAX_PATH];
    TCHAR    szCLSID[MAX_PATH];
    TCHAR    szModule[MAX_PATH];
    LPWSTR   pwsz;

    //get the CLSID in string form
    StringFromIID(clsid, &pwsz);

    if (pwsz)
    {
#ifdef UNICODE
        lstrcpy(szCLSID, pwsz);
#else
        WideCharToMultiByte( CP_ACP,
            0,
            pwsz,
            -1,
            szCLSID,
            ARRAYSIZE(szCLSID),
            NULL,
            NULL);
#endif

        //free the string
        LPMALLOC pMalloc;
        CoGetMalloc(1, &pMalloc);
        pMalloc->Free(pwsz);
        pMalloc->Release();
    }

    //get this app's path and file name
    GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule));

    DOREGSTRUCT ClsidEntries[] = {
        HKEY_CLASSES_ROOT,   TEXT("CLSID\\%s"),                  NULL,                   lpszTitle,
        HKEY_CLASSES_ROOT,   TEXT("CLSID\\%s\\InprocServer32"),  NULL,                   szModule,
        HKEY_CLASSES_ROOT,   TEXT("CLSID\\%s\\InprocServer32"),  TEXT("ThreadingModel"), TEXT("Apartment"),
        NULL,                NULL,                               NULL,                   NULL
    };

    //register the CLSID entries
    for(i = 0; ClsidEntries[i].hRootKey; i++)
    {
        //create the sub key string - for this case, insert the file extension
        _stprintf_s(szSubKey, ClsidEntries[i].szSubKey, szCLSID);

        lResult = RegCreateKeyEx(  ClsidEntries[i].hRootKey,
            szSubKey,
            0,
            NULL,
            REG_OPTION_NON_VOLATILE,
            KEY_WRITE,
            NULL,
            &hKey,
            &dwDisp);

        if (NOERROR == lResult)
        {
            TCHAR szData[MAX_PATH];

            //if necessary, create the value string
            _stprintf_s(szData, ClsidEntries[i].szData, szModule);

            lResult = RegSetValueEx(hKey,
                ClsidEntries[i].lpszValueName,
                0,
                REG_SZ,
                (LPBYTE)szData,
                (lstrlen(szData) + 1) * sizeof(TCHAR));

            RegCloseKey(hKey);
        }
        else
            return FALSE;
    }

    return TRUE;
}

typedef struct{
    HKEY    hRootKey;
    LPTSTR    szSubKey;//TCHAR szSubKey[MAX_PATH];
}    DOUNREGSTRUCT, *LPDOUNREGSTRUCT;

BOOL UnregisterServer(CLSID clsid)
{
    int      i;
    LRESULT  lResult;
    TCHAR    szSubKey[MAX_PATH];
    TCHAR    szCLSID[MAX_PATH];
    LPWSTR   pwsz;

    //get the CLSID in string form
    StringFromIID(clsid, &pwsz);

    if (pwsz)
    {
#ifdef UNICODE
        lstrcpy(szCLSID, pwsz);
#else
        WideCharToMultiByte( CP_ACP,
            0,
            pwsz,
            -1,
            szCLSID,
            ARRAYSIZE(szCLSID),
            NULL,
            NULL);
#endif

        //free the string
        LPMALLOC pMalloc;
        CoGetMalloc(1, &pMalloc);
        pMalloc->Free(pwsz);
        pMalloc->Release();
    }

    DOUNREGSTRUCT ClsidEntries[] = {
        HKEY_CLASSES_ROOT,   TEXT("CLSID\\%s\\InprocServer32"),
        HKEY_CLASSES_ROOT,   TEXT("CLSID\\%s"),
        NULL,                NULL
    };

    //register the CLSID entries
    for(i = 0; ClsidEntries[i].hRootKey; i++)
    {
        _stprintf_s(szSubKey, ClsidEntries[i].szSubKey, szCLSID);
        lResult = RegDeleteKey(ClsidEntries[i].hRootKey, szSubKey);
        //if (ERROR_SUCCESS != lResult)
        //    return FALSE;
    }

return TRUE;
}

BOOL RegisterComCat(CLSID clsid, CATID CatID)
{
    ICatRegister   *pcr;
    HRESULT        hr = S_OK ;

    CoInitialize(NULL);

    hr = CoCreateInstance(  CLSID_StdComponentCategoriesMgr,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ICatRegister,
        (LPVOID*)&pcr);

    if (SUCCEEDED(hr))
    {
        hr = pcr->RegisterClassImplCategories(clsid, 1, &CatID);
        pcr->Release();
    }

    CoUninitialize();

    return SUCCEEDED(hr);
}

BOOL UnregisterComCat(CLSID clsid, CATID CatID)
{
    ICatRegister   *pcr;
    HRESULT        hr = S_OK ;

    CoInitialize(NULL);

    hr = CoCreateInstance(  CLSID_StdComponentCategoriesMgr,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ICatRegister,
        (LPVOID*)&pcr);

    if (SUCCEEDED(hr))
    {
        hr = pcr->UnRegisterClassImplCategories(clsid, 1, &CatID);
        pcr->Release();
    }

    CoUninitialize();

    //return SUCCEEDED(hr);
    return TRUE;
}
