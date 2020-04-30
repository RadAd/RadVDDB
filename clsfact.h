#ifndef CLASSFACTORY_H
#define CLASSFACTORY_H

#include <windows.h>

class CClassFactory : public IClassFactory
{
protected:
    DWORD m_ObjRefCount;

public:
    CClassFactory(CLSID);
    ~CClassFactory();

    //IUnknown methods
    STDMETHODIMP QueryInterface(REFIID, LPVOID*) override;
    STDMETHODIMP_(DWORD) AddRef() override;
    STDMETHODIMP_(DWORD) Release() override;

    //IClassFactory methods
    STDMETHODIMP CreateInstance(LPUNKNOWN, REFIID, LPVOID*) override;
    STDMETHODIMP LockServer(BOOL) override;

private:
    CLSID m_clsidObject;
};

#endif    //CLASSFACTORY_H
