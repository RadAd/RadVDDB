#pragma once

#include "Win10Desktops.h"

class VirtualDesktopNotification : public IVirtualDesktopNotification
{
private:
    HWND _hWnd;
    ULONG _referenceCount;
public:
    VirtualDesktopNotification(HWND hWnd);

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) override;
    STDMETHODIMP_(DWORD) AddRef() override;
    STDMETHODIMP_(DWORD) STDMETHODCALLTYPE Release() override;

    STDMETHODIMP VirtualDesktopCreated(IVirtualDesktop* pDesktop) override;
    STDMETHODIMP VirtualDesktopDestroyBegin(IVirtualDesktop* pDesktopDestroyed, IVirtualDesktop* pDesktopFallback) override;
    STDMETHODIMP VirtualDesktopDestroyFailed(IVirtualDesktop* pDesktopDestroyed, IVirtualDesktop* pDesktopFallback) override;
    STDMETHODIMP VirtualDesktopDestroyed(IVirtualDesktop* pDesktopDestroyed, IVirtualDesktop* pDesktopFallback) override;
    STDMETHODIMP ViewVirtualDesktopChanged(IApplicationView* pView) override;
    STDMETHODIMP CurrentVirtualDesktopChanged(IVirtualDesktop* pDesktopOld, IVirtualDesktop* pDesktopNew) override;
};
