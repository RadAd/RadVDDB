#include "DeskBand.h"
#include "Guid.h"
#include "winerror.h"
#include "Globals.h"
#include "Win10Desktops.h"
#include "VDNotification.h"
#include "ComUtils.h"

#include <tchar.h>
#include <windowsx.h>
#include <winstring.h>

#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

// TODO
// Use top left of virtual screen --- const POINT dtorig = { GetSystemMetrics(SM_XVIRTUALSCREEN), GetSystemMetrics(SM_YVIRTUALSCREEN) };

#include <wrl.h>
#include <windows.ui.viewmanagement.h>

namespace abi_vm = ABI::Windows::UI::ViewManagement;
namespace wrl = Microsoft::WRL;
namespace wf = Windows::Foundation;

COLORREF GetAccentColor()
{
    // Error checking has been elided for expository purposes.
    wrl::ComPtr<abi_vm::IUISettings3> settings;
    wf::ActivateInstance(wrl::Wrappers::HStringReference(RuntimeClass_Windows_UI_ViewManagement_UISettings).Get(), &settings);
    ABI::Windows::UI::Color color;
    settings->GetColorValue(abi_vm::UIColorType::UIColorType_Accent, &color);
    // color.A, color.R, color.G, and color.B are the color channels.
    return RGB(color.R, color.G, color.B);
}

#define IDM_CREATE      0

#pragma optimize("", off)

HWND g_hWnd = NULL;

#define WM_WINEVENT (WM_USER + 12)

void CALLBACK WinEventProc(
    HWINEVENTHOOK /*hook*/,
    DWORD event,
    HWND hwnd,
    LONG idObject,
    LONG idChild,
    DWORD /*idEventThread*/,
    DWORD /*time*/)
{
    if (idObject == OBJID_WINDOW &&
        idChild == CHILDID_SELF)
    {
        if (event == EVENT_SYSTEM_FOREGROUND
            || event == EVENT_OBJECT_SHOW
            || event == EVENT_OBJECT_HIDE
            || event == EVENT_OBJECT_LOCATIONCHANGE)
        {
            PostMessage(g_hWnd, WM_WINEVENT, (WPARAM) hwnd, (LPARAM) event);
        }
    }
}

CDeskBand::CDeskBand(CLSID pClassID)
{
    m_pClassID = pClassID;

    m_pSite = NULL;

    m_hWnd = NULL;
    m_hwndParent = NULL;

    m_bFocus = FALSE;

    m_dwViewMode = 0;
    m_dwBandID = 0;

    m_ObjRefCount = 1;
    g_DllRefCount++;

    m_idVirtualDesktopNotification = 0;
}

CDeskBand::~CDeskBand()
{
    UnregisterNotify();

    if (m_hook != NULL)
    {
        UnhookWinEvent(m_hook);
        m_hook = NULL;
        g_hWnd = NULL;
    }

    g_DllRefCount--;
}

// IUnknown Implementation
STDMETHODIMP CDeskBand::QueryInterface(REFIID riid, LPVOID* ppReturn)
{
    *ppReturn = NULL;

    if (IsEqualIID(riid, IID_IUnknown))                         //IUnknown
        *ppReturn = this;
    else if (IsEqualIID(riid, IID_IOleWindow))                  //IOleWindow
        *ppReturn = static_cast<IOleWindow*>(this);
    else if (IsEqualIID(riid, IID_IDockingWindow))              //IDockingWindow
        *ppReturn = static_cast<IDockingWindow*>(this);
    else if (IsEqualIID(riid, IID_IObjectWithSite))             //IObjectWithSite
        *ppReturn = static_cast<IObjectWithSite*>(this);
    else if (IsEqualIID(riid, IID_IDeskBand))                   //IDeskBand
        *ppReturn = static_cast<IDeskBand*>(this);
    else if (IsEqualIID(riid, IID_IDeskBand2))                  //IDeskBand2
        *ppReturn = static_cast<IDeskBand2*>(this);
    else if (IsEqualIID(riid, IID_IPersist))                    //IPersist
        *ppReturn = static_cast<IPersist*>(this);
    else if (IsEqualIID(riid, IID_IPersistStream))              //IPersistStream
        *ppReturn = static_cast<IPersistStream*>(this);
    else if (IsEqualIID(riid, IID_IContextMenu))                //IContextMenu
        *ppReturn = static_cast<IContextMenu*>(this);

    if (*ppReturn)
    {
        (*(LPUNKNOWN*) ppReturn)->AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

STDMETHODIMP_(DWORD) CDeskBand::AddRef()
{
    return ++m_ObjRefCount;
}

STDMETHODIMP_(DWORD) CDeskBand::Release()
{
    if (--m_ObjRefCount == 0)
    {
        delete this;
        return 0;
    }

    return m_ObjRefCount;
}

// IOleWindow Implementation
STDMETHODIMP CDeskBand::GetWindow(HWND* phWnd)
{
    // Connect doesn't work when explorer starts up
    // so we try again here
    Connect();

    *phWnd = m_hWnd;

    return S_OK;
}

STDMETHODIMP CDeskBand::ContextSensitiveHelp(BOOL fEnterMode)
{
    return E_NOTIMPL;
}

// IDockingWindow Implementation
STDMETHODIMP CDeskBand::ShowDW(BOOL fShow)
{
    if (m_hWnd)
        ShowWindow(m_hWnd, fShow ? SW_SHOW : SW_HIDE);

    return S_OK;
}

STDMETHODIMP CDeskBand::CloseDW(DWORD dwReserved)
{
    ShowDW(FALSE);

    if (IsWindow(m_hWnd))
        DestroyWindow(m_hWnd);

    m_hWnd = NULL;
    return S_OK;
}

STDMETHODIMP CDeskBand::ResizeBorderDW(LPCRECT prcBorder, IUnknown* punkSite, BOOL fReserved)
{
    return E_NOTIMPL;
}

// IObjectWithSite implementations

STDMETHODIMP CDeskBand::SetSite(IUnknown* punkSite)
{
    m_pSite.Release();

    //If punkSite is not NULL, a new site is being set.
    if (punkSite)
    {
        //Get the parent window.
        m_hwndParent = NULL;

        CComPtr<IOleWindow>  pOleWindow;
        if (SUCCEEDED(punkSite->QueryInterface(IID_IOleWindow, (LPVOID*) &pOleWindow)))
            pOleWindow->GetWindow(&m_hwndParent);

        if (!m_hwndParent)
            return E_FAIL;

        if (!RegisterAndCreateWindow())
            return E_FAIL;

        UnregisterNotify();

        m_pNotify = new VirtualDesktopNotification(m_hWnd);
        if (!m_pDesktopNotificationService || FAILED(m_pDesktopNotificationService->Register(m_pNotify, &m_idVirtualDesktopNotification)))
            OutputDebugString(L"RADVDDB: Error register DesktopNotificationService\n");
            //return E_FAIL;

        //Get and keep the IInputObjectSite pointer.
        if (SUCCEEDED(punkSite->QueryInterface(IID_IInputObjectSite, (LPVOID*) &m_pSite)))
            return S_OK;

        return E_FAIL;
    }

    return S_OK;
}

STDMETHODIMP CDeskBand::GetSite(REFIID riid, LPVOID* ppvReturn)
{
    *ppvReturn = NULL;

    if (m_pSite)
        return m_pSite->QueryInterface(riid, ppvReturn);

    return E_FAIL;
}

// IDeskBand implementation

STDMETHODIMP CDeskBand::GetBandInfo(DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO* pdbi)
{
    if (pdbi)
    {
        m_dwBandID = dwBandID;
        m_dwViewMode = dwViewMode;

        if (pdbi->dwMask & DBIM_MINSIZE)
        {
            pdbi->ptMinSize.x = 100;
            pdbi->ptMinSize.y = 10;
        }

        if (pdbi->dwMask & DBIM_MAXSIZE)
        {
            pdbi->ptMaxSize.x = -1;
            pdbi->ptMaxSize.y = -1;
        }

        if (pdbi->dwMask & DBIM_INTEGRAL)
        {
            pdbi->ptIntegral.x = 1;
            pdbi->ptIntegral.y = 1;
        }

        if (pdbi->dwMask & DBIM_ACTUAL)
        {
            pdbi->ptActual.x = 100;
            pdbi->ptActual.y = 10;
        }

        if (pdbi->dwMask & DBIM_TITLE)
        {
            lstrcpyW(pdbi->wszTitle, TEXT("Rad Virtual Desktop"));
        }

        if (pdbi->dwMask & DBIM_MODEFLAGS)
        {
            pdbi->dwModeFlags = DBIMF_VARIABLEHEIGHT;
        }

#if 0
        if (pdbi->dwMask & DBIM_BKCOLOR)
        {
            //Use the default background color by removing this flag.
            pdbi->dwMask &= ~DBIM_BKCOLOR;
        }
#endif

        return S_OK;
    }

    return E_INVALIDARG;
}

// IDeskBand2 implementation

STDMETHODIMP CDeskBand::CanRenderComposited(BOOL* pfCanRenderComposited)
{
    *pfCanRenderComposited = TRUE;
    return S_OK;
}

STDMETHODIMP CDeskBand::SetCompositionState(BOOL fCompositionEnabled)
{
    return E_NOTIMPL;
}

STDMETHODIMP CDeskBand::GetCompositionState(BOOL* pfCompositionEnabled)
{
    return E_NOTIMPL;
}


// IPersistStream implementations
//
// This is only supported to allow the desk band to be dropped on the
// desktop and to prevent multiple instances of the desk band from showing
// up in the context menu. This desk band doesn't actually persist any data.

STDMETHODIMP CDeskBand::GetClassID(LPCLSID pClassID)
{
    *pClassID = m_pClassID;

    return S_OK;
}

STDMETHODIMP CDeskBand::IsDirty(void)
{
    return S_FALSE;
}

STDMETHODIMP CDeskBand::Load(LPSTREAM pStream)
{
    return S_OK;
}

STDMETHODIMP CDeskBand::Save(LPSTREAM pStream, BOOL fClearDirty)
{
    return S_OK;
}

STDMETHODIMP CDeskBand::GetSizeMax(ULARGE_INTEGER* pul)
{
    return E_NOTIMPL;
}

// IContextMenu Implementation

STDMETHODIMP CDeskBand::QueryContextMenu(HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    Connect();
    if (!(CMF_DEFAULTONLY & uFlags))
    {
        InsertMenu(hMenu, indexMenu++, MF_STRING | MF_BYPOSITION, idCmdFirst + IDM_CREATE, TEXT("&Create Desktop"));

        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(IDM_CREATE - IDM_CREATE + 1)); // Max - Min + 1
    }

    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));
}

STDMETHODIMP CDeskBand::InvokeCommand(LPCMINVOKECOMMANDINFO lpcmi)
{
    switch (LOWORD(lpcmi->lpVerb))
    {
    case IDM_CREATE:
    {
        CComPtr<Win10::IVirtualDesktop> pNewDesktop;
        if (!m_pDesktopManagerInternal || FAILED(m_pDesktopManagerInternal->CreateDesktopW(&pNewDesktop)))
            MessageBox(m_hWnd, L"Error creating virtual desktop", L"Error", MB_ICONERROR | MB_OK);
    }
    return NOERROR;

    default:
        return E_INVALIDARG;
    }

    return NOERROR;
}

STDMETHODIMP CDeskBand::GetCommandString(UINT_PTR idCommand, UINT uFlags, LPUINT lpReserved, LPSTR lpszName, UINT uMaxNameLen)
{
    HRESULT  hr = E_INVALIDARG;
    LPTSTR lptszName = (LPTSTR) lpszName;

    switch (uFlags)
    {
    case GCS_HELPTEXT:
        switch (idCommand)
        {
        case IDM_CREATE:
            _tcscpy_s(lptszName, uMaxNameLen, TEXT("Create a new virtual desktop"));
            hr = NOERROR;
            break;
        }
        break;

    case GCS_VERB:
        switch (idCommand)
        {
        case IDM_CREATE:
            _tcscpy_s(lptszName, uMaxNameLen, TEXT("create"));
            hr = NOERROR;
            break;
        }
        break;

    case GCS_VALIDATE:
        hr = NOERROR;
        break;
    }

    return hr;
}

// private method implementations

LRESULT CALLBACK CDeskBand::WndProc(HWND hWnd, UINT uMessage, WPARAM wParam, LPARAM lParam)
{
    CDeskBand* pThis = (CDeskBand*) GetWindowLongPtr(hWnd, GWLP_USERDATA);

    switch (uMessage)
    {
    case WM_NCCREATE:
    {
        LPCREATESTRUCT lpcs = (LPCREATESTRUCT) lParam;
        pThis = (CDeskBand*) (lpcs->lpCreateParams);
        SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR) pThis);

        pThis->m_hWnd = hWnd;
    }
    break;

    case WM_DESTROY:
        pThis->OnDestroy();
        break;

    case WM_PAINT:
        return pThis->OnPaint();

    case WM_LBUTTONDOWN:
        return pThis->OnLButtonDown((UINT) wParam, POINT({ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) }));

    case WM_LBUTTONUP:
        return pThis->OnLButtonUp((UINT) wParam, POINT({ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) }));

    case WM_SETFOCUS:
        return pThis->OnSetFocus();

    case WM_KILLFOCUS:
        return pThis->OnKillFocus();

    case WM_WINEVENT:
        pThis->OnWinEvent((HWND) wParam, (DWORD) lParam);
        return 0;
    }

    return DefWindowProc(hWnd, uMessage, wParam, lParam);
}

void CDeskBand::OnDestroy(void)
{
    UnregisterNotify();
}

LRESULT CDeskBand::OnPaint(void)
{
    PAINTSTRUCT ps;
    BeginPaint(m_hWnd, &ps);

    RECT        rc;
    GetClientRect(m_hWnd, &rc);

    const SIZE dtsz = { GetSystemMetrics(SM_CXVIRTUALSCREEN), GetSystemMetrics(SM_CYVIRTUALSCREEN) };

    HGDIOBJ    OrigPen = SelectObject(ps.hdc, GetStockObject(WHITE_PEN));
    HGDIOBJ    OrigBrush = SelectObject(ps.hdc, GetStockObject(WHITE_BRUSH));
    HGDIOBJ    OrigFont = SelectObject(ps.hdc, GetStockObject(DEVICE_DEFAULT_FONT));

    GDIPtr<HPEN> hPen(GetStockPen(BLACK_PEN));
    GDIPtr<HPEN> hSelectedPen(CreatePen(PS_SOLID, 5, GetAccentColor()));
    GDIPtr<HBRUSH> hBackgroundBrush(GetSysColorBrush(COLOR_WINDOWFRAME));
    GDIPtr<HBRUSH> hWindowBrush(GetSysColorBrush(COLOR_WINDOW));

    GDIPtr<HFONT> hTitleFont(CreateFont(-12, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, DEFAULT_PITCH, TEXT("Arial")));
    SetBkMode(ps.hdc, TRANSPARENT);

    CComPtr<Win10::IVirtualDesktop> pCurrentDesktop;
    if (!m_pDesktopManagerInternal || FAILED(m_pDesktopManagerInternal->GetCurrentDesktop(&pCurrentDesktop)))
        OutputDebugString(L"RADVDDB: Warning GetCurrentDesktop\n");

    LONG w;
    RECT dtrc = GetFirstDesktopRect(dtsz, w);
    const SIZE dtrcsz = GetSize(dtrc);

    CComPtr<IObjectArray> pViewArray;
    if (m_pApplicationViewCollection && SUCCEEDED(m_pApplicationViewCollection->GetViewsByZOrder(&pViewArray)))
        OutputDebugString(L"RADVDDB: Warning GetViewsByZOrder\n");

    CComPtr<IObjectArray> pDesktopArray;
    if (m_pDesktopManagerInternal && SUCCEEDED(m_pDesktopManagerInternal->GetDesktops(&pDesktopArray)))
    {
        for (CComPtr<Win10::IVirtualDesktop> pDesktop : ObjectArrayRange<Win10::IVirtualDesktop>(pDesktopArray))
        {
            SelectObject(ps.hdc, hPen);
            SelectObject(ps.hdc, hBackgroundBrush);
            Rectangle(ps.hdc, dtrc);

            if (pViewArray)
            {
                SelectObject(ps.hdc, hPen);
                SelectObject(ps.hdc, hWindowBrush);
                int s = SaveDC(ps.hdc);
                IntersectClipRect(ps.hdc, dtrc.left, dtrc.top, dtrc.right, dtrc.bottom);

                for (CComPtr<IApplicationView> pView : ObjectArrayRangeRev<IApplicationView>(pViewArray))
                {
                    BOOL bShowInSwitchers = FALSE;
                    if (SUCCEEDED(pView->GetShowInSwitchers(&bShowInSwitchers)))
                    {
                    }

                    if (!bShowInSwitchers)
                        continue;

                    BOOL bVisible = FALSE;
                    if (SUCCEEDED(pDesktop->IsViewVisible(pView, &bVisible)))
                    {
                    }

                    if (!bVisible)
                        continue;

                    HWND hWnd;
                    if (SUCCEEDED(pView->GetThumbnailWindow(&hWnd)))
                    {
                        RECT        wrc;
                        GetWindowRect(hWnd, &wrc);

                        Multiply(wrc, dtrcsz);
                        Divide(wrc, dtsz);
                        OffsetRect(&wrc, dtrc.left, dtrc.top);

                        Rectangle(ps.hdc, wrc);
                    }
                }

                CComPtr<Win10::IVirtualDesktop2> pDesktop2;
                if (SUCCEEDED(pDesktop.QueryInterface(&pDesktop2)))
                {
                    HSTRING name = NULL;
                    if (SUCCEEDED(pDesktop2->GetName(&name)))
                    {
                        SelectObject(ps.hdc, hTitleFont);

                        UINT32 length = 0;
                        PCWSTR pName = WindowsGetStringRawBuffer(name, &length);

#if 0
                        SetBkColor(ps.hdc, RGB(0, 0, 0));
                        SetTextColor(ps.hdc, RGB(185, 255, 70));
                        SetTextColor(ps.hdc, GetAccentColor());
                        DrawText(ps.hdc, pName, length, &dtrc, DT_LEFT | DT_TOP);
#else
                        DrawShadowText(
                            ps.hdc,
                            pName,
                            length,
                            &dtrc,
                            DT_LEFT | DT_TOP,
                            RGB(255, 255, 255),
                            RGB(0, 0, 0),
                            0,
                            0
                        );
#endif
                        WindowsDeleteString(name);
                    }
                }

                RestoreDC(ps.hdc, s);
            }

            if (pCurrentDesktop == pDesktop)
            {
                SelectObject(ps.hdc, hSelectedPen);
                MoveToEx(ps.hdc, dtrc.left, dtrc.bottom - 1, nullptr);
                LineTo(ps.hdc, dtrc.right - 1, dtrc.bottom - 1);
            }

            OffsetRect(&dtrc, w, 0);
        }
    }

    SelectObject(ps.hdc, OrigPen);
    SelectObject(ps.hdc, OrigBrush);
    SelectObject(ps.hdc, OrigFont);

    EndPaint(m_hWnd, &ps);

    return 0;
}

LRESULT CDeskBand::OnLButtonDown(UINT uModKeys, POINT pt)
{
    SetCapture(m_hWnd);
    pCaptureDesktop = GetDesktop(pt);
    // TODO Visual representation of click
    return 0;
}

LRESULT CDeskBand::OnLButtonUp(UINT uModKeys, POINT pt)
{
    ReleaseCapture();
    // TODO Check mouse didn't move too far

    CComPtr<Win10::IVirtualDesktop> pDesktop = GetDesktop(pt);

    if (pDesktop && pCaptureDesktop == pDesktop)
    {
        if (!m_pDesktopManagerInternal || FAILED(m_pDesktopManagerInternal->SwitchDesktop(pDesktop)))
            MessageBox(m_hWnd, L"Error switching desktops", L"Error", MB_ICONERROR | MB_OK);;
    }

    return 0;
}

void CDeskBand::OnWinEvent(HWND hWnd, DWORD event)
{
    // TODO Check in hWnd visible
    bool refresh = false;

    CComPtr<IApplicationView> pView;
    if (m_pApplicationViewCollection && SUCCEEDED(m_pApplicationViewCollection->GetViewForHwnd(hWnd, &pView)))
    {
        BOOL bShowInSwitchers = FALSE;
        if (SUCCEEDED(pView->GetShowInSwitchers(&bShowInSwitchers)))
        {
        }

        refresh = bShowInSwitchers != FALSE;
    }

    if (refresh)
        InvalidateRect(m_hWnd, nullptr, FALSE);
}

void CDeskBand::FocusChange(BOOL bFocus)
{
    m_bFocus = bFocus;

    //inform the input object site that the focus has changed
    if (m_pSite)
        m_pSite->OnFocusChangeIS((IDockingWindow*) this, bFocus);
}

LRESULT CDeskBand::OnSetFocus(void)
{
    FocusChange(TRUE);
    return 0;
}

LRESULT CDeskBand::OnKillFocus(void)
{
    FocusChange(FALSE);
    return 0;
}

BOOL CDeskBand::RegisterAndCreateWindow(void)
{
    //If the window doesn't exist yet, create it now.
    if (!m_hWnd)
    {
        //Can't create a child window without a parent.
        if (!m_hwndParent)
            return FALSE;

        //If the window class has not been registered, then do so.
        WNDCLASS wc;
        if (!GetClassInfo(g_hInst, DB_CLASS_NAME, &wc))
        {
            ZeroMemory(&wc, sizeof(wc));
            wc.style = CS_HREDRAW | CS_VREDRAW | CS_GLOBALCLASS;
            wc.lpfnWndProc = WndProc;
            wc.cbClsExtra = 0;
            wc.cbWndExtra = 0;
            wc.hInstance = g_hInst;
            wc.hIcon = NULL;
            wc.hCursor = LoadCursor(NULL, IDC_ARROW);
            wc.hbrBackground = (HBRUSH) GetStockObject(BLACK_BRUSH);
            wc.lpszMenuName = NULL;
            wc.lpszClassName = DB_CLASS_NAME;

            if (!RegisterClass(&wc))
            {
                //If RegisterClass fails, CreateWindow below will fail.
            }
        }

        RECT  rc;
        GetClientRect(m_hwndParent, &rc);

        //Create the window. The WndProc will set m_hWnd.
        CreateWindowEx(0,
            DB_CLASS_NAME,
            NULL,
            WS_CHILD | WS_CLIPSIBLINGS,
            rc.left,
            rc.top,
            rc.right - rc.left,
            rc.bottom - rc.top,
            m_hwndParent,
            NULL,
            g_hInst,
            (LPVOID) this);
    }

    return (NULL != m_hWnd);
}

void CDeskBand::Connect()
{
    if (!m_pDesktopManagerInternal || !m_pDesktopNotificationService)
    {
        CComPtr<IServiceProvider> pServiceProvider;
        pServiceProvider.CoCreateInstance(CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER);
        if (pServiceProvider)
        {
            if (!m_pDesktopManagerInternal)
            {
                if (FAILED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, &m_pDesktopManagerInternal)))
                    OutputDebugString(L"RADVDDB: Error obtaining CLSID_VirtualDesktopManagerInternal\n");
            }
            if (!m_pApplicationViewCollection)
            {
                if (FAILED(pServiceProvider->QueryService(__uuidof(m_pApplicationViewCollection), &m_pApplicationViewCollection)))
                    OutputDebugString(L"RADVDDB: Error obtaining IApplicationViewCollection\n");
            }
            if (!m_pDesktopNotificationService)
            {
                if (FAILED(pServiceProvider->QueryService(CLSID_VirtualNotificationService, &m_pDesktopNotificationService)))
                    OutputDebugString(L"RADVDDB: Error obtaining CLSID_VirtualNotificationService\n");
            }
        }

        if (m_hook == NULL && m_hWnd != NULL)
        {
            g_hWnd = m_hWnd;
            const DWORD processId = 0;
            const DWORD threadId = 0;
            m_hook = SetWinEventHook(
                EVENT_OBJECT_SHOW,
                EVENT_OBJECT_LOCATIONCHANGE,
                nullptr,
                WinEventProc,
                processId,
                threadId,
                WINEVENT_OUTOFCONTEXT);
            if (m_hook == NULL)
            {
                OutputDebugString(L"RADVDDB: Error register SetWinEventHook\n");
                g_hWnd = NULL;
            }
        }

        if (m_hWnd != NULL)
            InvalidateRect(m_hWnd, nullptr, FALSE);
    }

    if (m_pNotify && m_idVirtualDesktopNotification == 0)
    {
        if (!m_pDesktopNotificationService || FAILED(m_pDesktopNotificationService->Register(m_pNotify, &m_idVirtualDesktopNotification)))
            OutputDebugString(L"RADVDDB: Error register DesktopNotificationService\n");
    }
}

void CDeskBand::UnregisterNotify(void)
{
    if (m_idVirtualDesktopNotification != 0)
    {
        if (!m_pDesktopNotificationService || FAILED(m_pDesktopNotificationService->Unregister(m_idVirtualDesktopNotification)))
            OutputDebugString(L"RADVDDB: Error unregister notification\n");
        m_idVirtualDesktopNotification = 0;
    }
}

RECT CDeskBand::GetFirstDesktopRect(const SIZE dtsz, LONG& w)
{
    CComPtr<IObjectArray> pDesktopArray;
    if (!m_pDesktopManagerInternal || FAILED(m_pDesktopManagerInternal->GetDesktops(&pDesktopArray)))
        OutputDebugString(L"RADVDDB: Warning GetDesktops\n");

    UINT count = 1;
    if (!pDesktopArray || FAILED(pDesktopArray->GetCount(&count)))
        OutputDebugString(L"RADVDDB: Warning GetCount\n");;

    RECT rc;
    GetClientRect(m_hWnd, &rc);

    const double dtar = AspectRatio(dtsz);

    w = (rc.right - rc.left + 5) / count;
    RECT vdrc(rc);
    vdrc.right = vdrc.left + w - 5;
    InflateRect(&vdrc, 0, -5);

    const SIZE vdsz = GetSize(vdrc);
    const LONG vdnh = (LONG) (vdsz.cx / dtar + 0.5);
    vdrc.top = vdrc.top + (vdsz.cy - vdnh) / 2;
    vdrc.bottom = vdrc.top + vdnh;

    const double vdar = AspectRatio(GetSize(vdrc));

    return vdrc;
}

CComPtr<Win10::IVirtualDesktop> CDeskBand::GetDesktop(POINT pt)
{
    const SIZE dtsz = { GetSystemMetrics(SM_CXVIRTUALSCREEN), GetSystemMetrics(SM_CYVIRTUALSCREEN) };

    LONG w;
    RECT dtrc = GetFirstDesktopRect(dtsz, w);

    CComPtr<IObjectArray> pObjectArray;
    if (!m_pDesktopManagerInternal || FAILED(m_pDesktopManagerInternal->GetDesktops(&pObjectArray)))
        return nullptr;
    for (CComPtr<Win10::IVirtualDesktop> pDesktop : ObjectArrayRange<Win10::IVirtualDesktop>(pObjectArray))
    {
        if (PtInRect(&dtrc, pt))
            return pDesktop;

        OffsetRect(&dtrc, w, 0);
    }

    return nullptr;
}
