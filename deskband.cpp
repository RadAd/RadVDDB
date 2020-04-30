#include "DeskBand.h"
#include "Guid.h"
#include "winerror.h"
#include "Globals.h"
#include "Win10Desktops.h"
#include "VDNotification.h"
#include "ComUtils.h"

#include <tchar.h>
#include <windowsx.h>

#define IDM_CREATE      0

#pragma optimize("", off)

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

    CComPtr<IServiceProvider> pServiceProvider;
    pServiceProvider.CoCreateInstance(CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER);
    if (FAILED(pServiceProvider->QueryService(CLSID_VirtualDesktopManagerInternal, IID_PPV_ARGS(&m_pDesktopManagerInternal))))
        MessageBox(m_hWnd, L"Error obtaining CLSID_VirtualDesktopManagerInternal", L"Error", MB_ICONERROR | MB_OK);
    if (FAILED(pServiceProvider->QueryService(CLSID_IVirtualNotificationService, IID_PPV_ARGS(&m_pDesktopNotificationService))))
        MessageBox(m_hWnd, L"Error obtaining CLSID_IVirtualNotificationService", L"Error", MB_ICONERROR | MB_OK);
    m_idVirtualDesktopNotification = 0;
}

CDeskBand::~CDeskBand()
{
    UnregisterNotify();

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
        if (FAILED(m_pDesktopNotificationService->Register(m_pNotify, &m_idVirtualDesktopNotification)))
            return E_FAIL;

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
        CComPtr<IVirtualDesktop> pNewDesktop;
        if (FAILED(m_pDesktopManagerInternal->CreateDesktopW(&pNewDesktop)))
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

    HGDIOBJ    OrigPen = SelectObject(ps.hdc, GetStockObject(WHITE_PEN));
    HGDIOBJ    OrigBrush = SelectObject(ps.hdc, GetStockObject(WHITE_BRUSH));

    GDIPtr<HPEN> hBorderPen((HPEN) GetStockObject(BLACK_PEN));
    GDIPtr<HBRUSH> hBorderBrush(GetSysColorBrush(COLOR_WINDOWFRAME));

    GDIPtr<HPEN> hBorderSelectedPen((HPEN) GetStockObject(BLACK_PEN));
    GDIPtr<HBRUSH> hBorderSelectedBrush(GetSysColorBrush(COLOR_WINDOW));

    CComPtr<IVirtualDesktop> pCurrentDesktop;
    if (FAILED(m_pDesktopManagerInternal->GetCurrentDesktop(&pCurrentDesktop)))
        ;

    LONG w;
    RECT dtrc = GetFirstDesktopRect(w);

    CComPtr<IObjectArray> pObjectArray;
    if (SUCCEEDED(m_pDesktopManagerInternal->GetDesktops(&pObjectArray)))
    {
        for (CComPtr<IVirtualDesktop> pDesktop : ObjectArrayRange<IVirtualDesktop>(pObjectArray))
        {
            SelectObject(ps.hdc, pCurrentDesktop == pDesktop ? hBorderSelectedPen : hBorderPen);
            SelectObject(ps.hdc, pCurrentDesktop == pDesktop ? hBorderSelectedBrush : hBorderBrush);
            Rectangle(ps.hdc, dtrc.left, dtrc.top, dtrc.right, dtrc.bottom);

            OffsetRect(&dtrc, w, 0);
        }
    }

    SelectObject(ps.hdc, OrigPen);
    SelectObject(ps.hdc, OrigBrush);

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

    CComPtr<IVirtualDesktop> pDesktop = GetDesktop(pt);

    if (pDesktop && pCaptureDesktop == pDesktop)
    {
        if (FAILED(m_pDesktopManagerInternal->SwitchDesktop(pDesktop)))
            MessageBox(m_hWnd, L"Error switching desktops", L"Error", MB_ICONERROR | MB_OK);;
    }

    return 0;
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

void CDeskBand::UnregisterNotify(void)
{
    if (m_idVirtualDesktopNotification != 0)
    {
        if (FAILED(m_pDesktopNotificationService->Unregister(m_idVirtualDesktopNotification)))
            MessageBox(m_hWnd, L"Error unregister notification", L"Error", MB_ICONERROR | MB_OK);
        m_idVirtualDesktopNotification = 0;
    }
}

RECT CDeskBand::GetFirstDesktopRect(LONG& w)
{
    CComPtr<IObjectArray> pObjectArray;
    if (FAILED(m_pDesktopManagerInternal->GetDesktops(&pObjectArray)))
        ;

    UINT count;
    if (FAILED(pObjectArray->GetCount(&count)))
        ;

    RECT rc;
    GetClientRect(m_hWnd, &rc);

    const SIZE dtsz = { GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN) };
    const double dtar = (double) dtsz.cx / dtsz.cy;

    w = (rc.right - rc.left + 5) / count;
    RECT vdrc(rc);
    vdrc.right = vdrc.left + w - 5;
    InflateRect(&vdrc, 0, -5);

    const LONG vdw = (vdrc.right - vdrc.left);
    const LONG vdh = (vdrc.bottom - vdrc.top);
    const LONG vdnh = (LONG) (vdw / dtar + 0.5);
    vdrc.top = vdrc.top + (vdh - vdnh) / 2;
    vdrc.bottom = vdrc.top + vdnh;

    const double vdar = ((double) vdrc.right - vdrc.left) / ((double) vdrc.bottom - vdrc.top);

    return vdrc;
}

CComPtr<IVirtualDesktop> CDeskBand::GetDesktop(POINT pt)
{
    LONG w;
    RECT dtrc = GetFirstDesktopRect(w);

    CComPtr<IObjectArray> pObjectArray;
    if (FAILED(m_pDesktopManagerInternal->GetDesktops(&pObjectArray)))
        return nullptr;
    for (CComPtr<IVirtualDesktop> pDesktop : ObjectArrayRange<IVirtualDesktop>(pObjectArray))
    {
        if (PtInRect(&dtrc, pt))
            return pDesktop;

        OffsetRect(&dtrc, w, 0);
    }

    return nullptr;
}
