#ifndef DESKBAND_H
#define DESKBAND_H

#include <windows.h>
#include <atlbase.h>
#include <shlobj.h>

#include "GdiUtils.h"

#define DB_CLASS_NAME (TEXT("RadVDDBClass"))

interface IVirtualDesktopManagerInternal;
interface IApplicationViewCollection;
interface IVirtualDesktopNotificationService;
interface IVirtualDesktopNotification;
interface IVirtualDesktop;

/**************************************************************************

   CDeskBand class definition

**************************************************************************/

class CDeskBand : public IDeskBand2,
                  public IObjectWithSite,
                  public IPersistStream,
                  public IContextMenu
{
protected:
    DWORD m_ObjRefCount;

public:
    CDeskBand(CLSID pClassID);
    ~CDeskBand();

    //IUnknown methods
    STDMETHODIMP QueryInterface(REFIID, LPVOID*) override;
    STDMETHODIMP_(DWORD) AddRef() override;
    STDMETHODIMP_(DWORD) Release() override;

    //IOleWindow methods
    STDMETHOD (GetWindow) (HWND*) override;
    STDMETHOD (ContextSensitiveHelp) (BOOL) override;

    //IDockingWindow methods
    STDMETHOD (ShowDW) (BOOL fShow) override;
    STDMETHOD (CloseDW) (DWORD dwReserved) override;
    STDMETHOD (ResizeBorderDW) (LPCRECT prcBorder, IUnknown* punkToolbarSite, BOOL fReserved) override;

    //IDeskBand methods
    STDMETHOD (GetBandInfo) (DWORD, DWORD, DESKBANDINFO*) override;

    //IDeskBand2 methods
    STDMETHOD(CanRenderComposited) (BOOL*) override;
    STDMETHOD(SetCompositionState) (BOOL) override;
    STDMETHOD(GetCompositionState) (BOOL*) override;

    //IObjectWithSite methods
    STDMETHOD (SetSite) (IUnknown*) override;
    STDMETHOD (GetSite) (REFIID, LPVOID*) override;

    //IPersistStream methods
    STDMETHOD (GetClassID) (LPCLSID) override;
    STDMETHOD (IsDirty) (void) override;
    STDMETHOD (Load) (LPSTREAM) override;
    STDMETHOD (Save) (LPSTREAM, BOOL) override;
    STDMETHOD (GetSizeMax) (ULARGE_INTEGER*) override;

    //IContextMenu methods
    STDMETHOD (QueryContextMenu)(HMENU, UINT, UINT, UINT, UINT) override;
    STDMETHOD (InvokeCommand)(LPCMINVOKECOMMANDINFO) override;
    STDMETHOD (GetCommandString)(UINT_PTR, UINT, LPUINT, LPSTR, UINT) override;

private:
    CLSID m_pClassID;

    BOOL m_bFocus;
    HWND m_hwndParent;
    HWND m_hWnd;
    DWORD m_dwViewMode;
    DWORD m_dwBandID;
    CComPtr<IInputObjectSite> m_pSite;

    CComPtr<IVirtualDesktopManagerInternal> m_pDesktopManagerInternal;
    CComPtr<IApplicationViewCollection> m_pApplicationViewCollection;
    CComPtr<IVirtualDesktopNotificationService> m_pDesktopNotificationService;
    DWORD m_idVirtualDesktopNotification = 0;
    CComPtr<IVirtualDesktopNotification> m_pNotify = nullptr;
    CComPtr<IVirtualDesktop> pCaptureDesktop;

private:
    void FocusChange(BOOL);
    LRESULT OnKillFocus(void);
    LRESULT OnSetFocus(void);
    static LRESULT CALLBACK WndProc(HWND hWnd, UINT uMessage, WPARAM wParam, LPARAM lParam);
    void OnDestroy(void);
    LRESULT OnPaint(void);
    LRESULT OnLButtonDown(UINT uModKeys, POINT pt);
    LRESULT OnLButtonUp(UINT uModKeys, POINT pt);
    BOOL RegisterAndCreateWindow(void);

    void Connect();
    void UnregisterNotify(void);
    RECT GetFirstDesktopRect(const SIZE dtsz, LONG& w);
    CComPtr<IVirtualDesktop> GetDesktop(POINT pt);
};

#endif   //DESKBAND_H

