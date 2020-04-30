#include "VDNotification.h"

VirtualDesktopNotification::VirtualDesktopNotification(HWND hWnd)
    : _hWnd(hWnd), _referenceCount(0)
{
}

STDMETHODIMP VirtualDesktopNotification::QueryInterface(REFIID riid, void** ppvObject)
{
    // Always set out parameter to NULL, validating it first.
    if (!ppvObject)
        return E_INVALIDARG;
    *ppvObject = NULL;

    if (riid == IID_IUnknown || riid == IID_IVirtualDesktopNotification)
    {
        // Increment the reference count and return the pointer.
        *ppvObject = (LPVOID) this;
        AddRef();
        return S_OK;
    }
    return E_NOINTERFACE;
}

STDMETHODIMP_(DWORD) VirtualDesktopNotification::AddRef()
{
    return InterlockedIncrement(&_referenceCount);
}

STDMETHODIMP_(DWORD) STDMETHODCALLTYPE VirtualDesktopNotification::Release()
{
    ULONG result = InterlockedDecrement(&_referenceCount);
    if (result == 0)
    {
        delete this;
    }
    return 0;
}

STDMETHODIMP VirtualDesktopNotification::VirtualDesktopCreated(IVirtualDesktop* pDesktop)
{
    InvalidateRect(_hWnd, nullptr, TRUE);
    return S_OK;
}

STDMETHODIMP VirtualDesktopNotification::VirtualDesktopDestroyBegin(IVirtualDesktop* pDesktopDestroyed, IVirtualDesktop* pDesktopFallback)
{
    return S_OK;
}

STDMETHODIMP VirtualDesktopNotification::VirtualDesktopDestroyFailed(IVirtualDesktop* pDesktopDestroyed, IVirtualDesktop* pDesktopFallback)
{
    return S_OK;
}

STDMETHODIMP VirtualDesktopNotification::VirtualDesktopDestroyed(IVirtualDesktop* pDesktopDestroyed, IVirtualDesktop* pDesktopFallback)
{
    InvalidateRect(_hWnd, nullptr, TRUE);
    return S_OK;
}

STDMETHODIMP VirtualDesktopNotification::ViewVirtualDesktopChanged(IApplicationView* pView)
{
    InvalidateRect(_hWnd, nullptr, TRUE);
    return S_OK;
}

STDMETHODIMP VirtualDesktopNotification::CurrentVirtualDesktopChanged(IVirtualDesktop* pDesktopOld, IVirtualDesktop* pDesktopNew)
{
    //viewCollection->RefreshCollection();
    InvalidateRect(_hWnd, nullptr, TRUE);
    return S_OK;
}
