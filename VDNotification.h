#pragma once

#include "Win10Desktops.h"

class VirtualDesktopNotification : public Win10::IVirtualDesktopNotification
{
private:
    HWND _hWnd;
    ULONG _referenceCount;
public:
    VirtualDesktopNotification(HWND hWnd);

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) override;
    STDMETHODIMP_(DWORD) AddRef() override;
    STDMETHODIMP_(DWORD) STDMETHODCALLTYPE Release() override;

    STDMETHODIMP VirtualDesktopCreated(Win10::IVirtualDesktop* pDesktop) override;
    STDMETHODIMP VirtualDesktopDestroyBegin(Win10::IVirtualDesktop* pDesktopDestroyed, Win10::IVirtualDesktop* pDesktopFallback) override;
    STDMETHODIMP VirtualDesktopDestroyFailed(Win10::IVirtualDesktop* pDesktopDestroyed, Win10::IVirtualDesktop* pDesktopFallback) override;
    STDMETHODIMP VirtualDesktopDestroyed(Win10::IVirtualDesktop* pDesktopDestroyed, Win10::IVirtualDesktop* pDesktopFallback) override;
    STDMETHODIMP ViewVirtualDesktopChanged(IApplicationView* pView) override;
    STDMETHODIMP CurrentVirtualDesktopChanged(Win10::IVirtualDesktop* pDesktopOld, Win10::IVirtualDesktop* pDesktopNew) override;
};
