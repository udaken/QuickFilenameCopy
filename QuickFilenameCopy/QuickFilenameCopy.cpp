#include "framework.h"
#include "resource.h"

#include <unordered_map>
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/result.h>

#pragma comment(lib, "Version.lib")
#pragma comment(lib, "Shlwapi.lib")

#pragma comment(linker, "/manifestdependency:\"type='win32' \
    name='Microsoft.Windows.Common-Controls' \
    version='6.0.0.0' \
    processorArchitecture='*' \
    publicKeyToken='6595b64144ccf1df' \
    language='*'\"")

#define MAX_LOADSTRING 100
#define WM_NOTIFYICON (WM_USER + 100)

#define CATCH_SHOW_MSGBOX()                                                         \
    catch (const wil::ResultException &e)                                           \
    {                                                                               \
        wchar_t message[2048]{};                                                    \
        wil::GetFailureLogString(message, ARRAYSIZE(message), e.GetFailureInfo());  \
        MessageBox(hWnd, message, g_szTitle, MB_ICONERROR);                         \
    }

#define TRACE() my::DbgPrint("{}", __FUNCTION__ "\n")
//#define TRACE() 
#define DBGPRINTLN(fmt, ...) my::DbgPrint(__FILE__ "(" _STRINGIZE(__LINE__) "): " fmt "\n", __VA_ARGS__)

using namespace std::string_literals;
using namespace std::string_view_literals;

#ifndef __cpp_lib_format
namespace std {
    using namespace ::fmt;
}
#endif

namespace my
{
    template <class... T>
    inline auto DbgPrint(const std::wstring_view fmt, const T&... args)
    {
        OutputDebugStringW(std::format(fmt, args...).c_str());
    }

    template <class... T>
    inline auto DbgPrint(const std::string_view fmt, const T&... args)
    {
        OutputDebugStringA(std::format(fmt, args...).c_str());
    }

    inline auto GetWindowThreadProcessId(HWND hwnd)
    {
        DWORD pid{};
        return std::make_tuple(::GetWindowThreadProcessId(hwnd, &pid), pid);
    }
    inline std::wstring GetClassName(HWND hwnd)
    {
        WCHAR className[256]{};
        unsigned len = ::GetClassName(hwnd, className, ARRAYSIZE(className));
        return { className, len };
    }
    inline std::wstring GetModuleFileName()
    {
        WCHAR fileName[256]{};
        unsigned len = ::GetModuleFileName(nullptr, fileName, ARRAYSIZE(fileName));
        return { fileName, len };
    }
} // namespace my

HINSTANCE g_hInst;
HHOOK g_hook;
WCHAR g_szTitle[MAX_LOADSTRING];
HWND g_hwnd;
wil::com_ptr_t<IShellWindows> g_pSHWinds;

LRESULT CALLBACK wndProc(HWND, UINT, WPARAM, LPARAM) noexcept;
INT_PTR CALLBACK about(HWND, UINT, WPARAM, LPARAM) noexcept;

constexpr auto NOTIFY_UID = 1;
constexpr auto szWindowClass = L"{1D93FDAB-20F9-427D-9650-8B9C861C8137}";
constexpr auto targetClassName{ L"CabinetWClass" };

template <typename err_policy = wil::err_exception_policy>
struct service_provider_t : wil::com_ptr_t<IServiceProvider, err_policy>
{
    //using namespace wil;
    using base = wil::com_ptr_t<IServiceProvider, err_policy>;
    using result = err_policy::result;

    template <typename E>
    explicit service_provider_t(wil::com_ptr_t<IServiceProvider, E> x)
        : base(x)
    {}

    template <class U, class E = err_policy>
    inline wil::com_ptr_t<U, E> query_service() const
    {
        static_assert(wistd::is_same<void, result>::value, "query requires exceptions or fail fast; use try_query or query_to");
        wil::com_ptr_t<U, E> s;

        err_policy::HResult(this->QueryService(IID_PPV_ARGS(&s)));
        return s;
    }
};

template <class T, typename err_policy = wil::err_exception_policy>
service_provider_t< err_policy> to_service_provider(wil::com_ptr_t<T> from)
{
    return service_provider_t{ from.query<IServiceProvider>() };
}

void setClipboardText(std::wstring_view ss)
{
    if (ss.empty())
        return;

    size_t cchNeeded = ss.size() + 1;
    wil::unique_hglobal hGlobal{ GlobalAlloc(GMEM_MOVEABLE | GMEM_DDESHARE, cchNeeded * sizeof(WCHAR)) };
    THROW_LAST_ERROR_IF_NULL(hGlobal);
    {
        WCHAR* lock = (reinterpret_cast<WCHAR*>(GlobalLock(hGlobal.get())));
        THROW_LAST_ERROR_IF_NULL(lock);
        THROW_IF_FAILED(StringCbCopyW(lock, GlobalSize(hGlobal.get()), ss.data()));
        GlobalUnlock(hGlobal.get());
    }

    THROW_IF_WIN32_BOOL_FALSE(OpenClipboard(g_hwnd));
    {
        auto defer = wil::scope_exit([&] {
            CloseClipboard();
        });
        THROW_IF_WIN32_BOOL_FALSE(EmptyClipboard());
        THROW_LAST_ERROR_IF_NULL(SetClipboardData(CF_UNICODETEXT, hGlobal.get()));
        auto _ = hGlobal.release();
    }
}

void copySelectedItems(wil::com_ptr_t<IFolderView2> pfv2)
{
    std::wstring ss;

    wil::com_ptr_t<IShellItemArray> pSIA;
    THROW_IF_FAILED(pfv2->GetSelection(TRUE, &pSIA));

    DWORD dwCount{};
    THROW_IF_FAILED(pSIA->GetCount(&dwCount));
    for (DWORD i = 0; i < dwCount; i++)
    {
        wil::com_ptr_t<IShellItem> pSI;
        THROW_IF_FAILED(pSIA->GetItemAt(i, &pSI));

        wil::unique_cotaskmem_string displayName;
        THROW_IF_FAILED(pSI->GetDisplayName(SIGDN_NORMALDISPLAY, &displayName));

        if (i > 0) {
            ss.append(L"\n");
        }

        ss.append(displayName.get());
    }
    setClipboardText(ss);

}

// Querying information from an Explorer window | The Old New Thing
// https://devblogs.microsoft.com/oldnewthing/20040720-00/?p=38393
void traverseShellWindows(wil::com_ptr_t<IShellWindows> pSHWinds, HWND hWndTarget)
{
    TRACE();

    if (hWndTarget == nullptr) {
        return;
    }

    long count;
    THROW_IF_FAILED(pSHWinds->get_Count(&count));
    for (long i = 0; i < count; i++)
    {
        VARIANT v{};
        V_VT(&v) = VT_I4; V_I4(&v) = i;
        wil::com_ptr_t< IDispatch> pDisp;
        THROW_IF_FAILED(pSHWinds->Item(v, pDisp.put()));
        auto pWBA = pDisp.query<IWebBrowserApp>();

        SHANDLE_PTR hWndShell{};
        THROW_IF_FAILED(pWBA->get_HWND(&hWndShell));

        { // debug
            wil::unique_bstr name;
            pWBA->get_Name(wil::out_param(name));

            wil::unique_bstr fullname;
            pWBA->get_FullName(wil::out_param(fullname));

            my::DbgPrint(L"[{}] name={}, fullname={}, hwnd={p}\n", i, name.get(), fullname.get(), hWndShell);
        }

        if (reinterpret_cast<HWND>(hWndShell) == hWndTarget)
        {
            auto psp = pWBA.query<IServiceProvider>();

            wil::com_ptr_t<IShellBrowser> psb;
            THROW_IF_FAILED(psp->QueryService(SID_STopLevelBrowser, __uuidof(IShellBrowser), psb.put_void()));

            wil::com_ptr_t<IShellView> psv;
            THROW_IF_FAILED(psb->QueryActiveShellView(&psv));

            auto pfv2 = psv.query<IFolderView2>();

            { // debug

                UINT viewMode{};
                THROW_IF_FAILED(pfv2->GetCurrentViewMode(&viewMode));

                int itemCount{};
                THROW_IF_FAILED(pfv2->ItemCount(SVGIO_BACKGROUND, &itemCount));

                my::DbgPrint(L"[{}] viewMode={}, itemCount={}\n", i, viewMode, itemCount);
            }

            copySelectedItems(pfv2);
            my::DbgPrint(L"[{}] Copied!\n", i);
            break;
        }
    }
}

LRESULT CALLBACK lowLevelKeyboardProc(int code, WPARAM wParam, LPARAM lParam) noexcept
{
    TRACE();

    if (code < HC_ACTION) return CallNextHookEx(nullptr, code, wParam, lParam);

    auto pKbdll = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);

    DBGPRINTLN("flags:{}, vkCode:{}", pKbdll->flags, pKbdll->vkCode);

    if ((pKbdll->flags & LLKHF_LOWER_IL_INJECTED) == 0 && pKbdll->vkCode == 'C')
    {
        auto ctrlKey = GetAsyncKeyState(VK_CONTROL) & 0x8000;
        auto shiftKey = GetAsyncKeyState(VK_SHIFT) & 0x8000;
        DBGPRINTLN("ctrlKey:{}, shiftKey:{}", ctrlKey, shiftKey);

        if (ctrlKey != 0 && shiftKey != 0)
        {
            auto hWnd{ GetFocus() };
            if (hWnd == nullptr) {
                hWnd = GetForegroundWindow();
            }

            if (hWnd != nullptr)
            {
                WCHAR className[512]{};
                RealGetWindowClass(hWnd, className, ARRAYSIZE(className));

                if (lstrcmp(className, targetClassName) == 0)
                {
                    traverseShellWindows(g_pSHWinds, hWnd);
                    return 1;
                }
            }
        }
    }

    return CallNextHookEx(nullptr, code, wParam, lParam);
}

void uninstallHook()
{
    if (g_hook != nullptr)
    {
        UnhookWindowsHookEx(g_hook);
        g_hook = nullptr;
    }
}

void installHook()
{
    TRACE();

    uninstallHook();

    DWORD dwTid = 0;/*global*/

    g_hook = SetWindowsHookEx(WH_KEYBOARD_LL, &lowLevelKeyboardProc, nullptr, dwTid);
    THROW_LAST_ERROR_IF_NULL(g_hook);
}

bool isWorking()
{
    return g_hook != nullptr;
}

BOOL addNotifyIcon(HWND hWnd, unsigned int uID)
{
    TRACE();
    NOTIFYICONDATA nid{ sizeof(nid) };

    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.hWnd = hWnd;
    nid.uID = uID;
    nid.uCallbackMessage = WM_NOTIFYICON;
    nid.hIcon = LoadIcon(g_hInst, MAKEINTRESOURCE(IDI_SMALL));
    wcscpy_s(nid.szTip, g_szTitle);

    return Shell_NotifyIcon(NIM_ADD, &nid);
}

void deleteNotifyIcon(HWND hWnd, unsigned int uID)
{
    TRACE();
    NOTIFYICONDATA nid = {};

    nid.cbSize = sizeof(nid);
    nid.hWnd = hWnd;
    nid.uID = uID;

    Shell_NotifyIcon(NIM_DELETE, &nid);
}

ATOM registerMyClass(HINSTANCE hInstance)
{
    TRACE();
    WNDCLASSEXW wcex{
        sizeof(WNDCLASSEX),
        CS_HREDRAW | CS_VREDRAW,
        &wndProc,
    };

    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL));
    wcex.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL));
    wcex.lpszClassName = szWindowClass;

    return RegisterClassExW(&wcex);
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine,
                      _In_ int nCmdShow)
{
    SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_SYSTEM32);

    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);
    UNREFERENCED_PARAMETER(nCmdShow);

    g_hInst = hInstance;

    LoadStringW(hInstance, IDS_APP_TITLE, g_szTitle, MAX_LOADSTRING);

    wcscat_s(g_szTitle,
#if _M_ARM64 
             L" (ARM64bit)"
#elif _M_X64
             L" (64bit)"
#elif _M_IX86 
             L" (32bit)"
#endif
    );

    wil::unique_mutex m{ CreateMutex(nullptr, FALSE, g_szTitle) };
    if (GetLastError() == ERROR_ALREADY_EXISTS)
    {
        ::MessageBox(nullptr, L"Another instance is already running.", g_szTitle, MB_ICONEXCLAMATION);
        return 0;
    }
    auto hr{ CoInitializeEx(NULL, COINIT_MULTITHREADED | COINIT_DISABLE_OLE1DDE) };
    //auto hr{ CoInitialize(nullptr) };
    THROW_IF_FAILED(hr);

    g_pSHWinds = wil::CoCreateInstance<ShellWindows, IShellWindows>(CLSCTX_ALL);
    //traverseShellWindows(g_pSHWinds, FindWindow(targetClassName, nullptr));

    installHook();

    registerMyClass(hInstance);

    HWND hWnd = CreateWindowW(szWindowClass, g_szTitle, 0, CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr,
                              hInstance, nullptr);

    if (hWnd == nullptr)
    {
        return FALSE;
    }

    if (!addNotifyIcon(hWnd, NOTIFY_UID))
    {
        return FALSE;
    }

    MSG msg{};
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    uninstallHook();
    deleteNotifyIcon(hWnd, NOTIFY_UID);

    return (int)msg.wParam;
}

void registerToShortcut(HWND hWnd)
try
{
    WCHAR targetPath[MAX_PATH]{};
    THROW_IF_WIN32_BOOL_FALSE(SHGetSpecialFolderPath(hWnd, targetPath, CSIDL_STARTUP, FALSE));
    StringCchCat(targetPath, ARRAYSIZE(targetPath), (L"\\"s + g_szTitle + L".lnk"s).c_str());

    auto pShellLink{ wil::CoCreateInstance<IShellLink>(CLSID_ShellLink) };
    auto pPersistFile{ pShellLink.query<IPersistFile>() };

    THROW_IF_FAILED(pShellLink->SetPath(my::GetModuleFileName().c_str()));
    THROW_IF_FAILED(pPersistFile->Save(targetPath, TRUE));
}
CATCH_SHOW_MSGBOX()

LRESULT CALLBACK wndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    static UINT s_uTaskbarRestart;
    static HMENU s_menu = nullptr;

    switch (message)
    {
    case WM_CREATE:
        g_hwnd = hWnd;
        s_uTaskbarRestart = RegisterWindowMessage(L"TaskbarCreated");
        s_menu = LoadMenu(g_hInst, MAKEINTRESOURCE(IDR_MENU1));
        return DefWindowProc(hWnd, message, wParam, lParam);
    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case ID_ROOT_ABOUT:
            DialogBox(g_hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, &about);
            break;
        case ID_ROOT_REGISTERTOSTARTUPPROGRAM:
            registerToShortcut(hWnd);
            break;
        case ID_ROOT_EXIT:
            DestroyWindow(hWnd);
            break;
        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
        }
        return 0;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    case WM_NOTIFYICON:
        switch (lParam)
        {
        case WM_RBUTTONDOWN:
        {
            POINT pt{};
            GetCursorPos(&pt);
            SetForegroundWindow(hWnd);
            TrackPopupMenu(GetSubMenu(s_menu, 0), TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, NULL);
        }
        break;
        default:
            ;
        }
        return 0;
    default:
        if (message == s_uTaskbarRestart)
            addNotifyIcon(hWnd, NOTIFY_UID);
        else
            return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

void getProductAndVersion(LPWSTR szCopyright, UINT uCopyrightLen, LPWSTR szProductVersion, UINT uProductVersionLen)
{
    WCHAR szFilename[MAX_PATH]{};
    GetModuleFileName(nullptr, szFilename, ARRAYSIZE(szFilename));

    DWORD dwHandle;
    DWORD dwSize = GetFileVersionInfoSize(szFilename, &dwHandle);
    if (dwSize == 0)
        std::abort();

    std::vector<BYTE> data(dwSize);
    if (!GetFileVersionInfo(szFilename, 0, dwSize, data.data()))
        std::abort();

    LPWSTR pvCopyright{}, pvProductVersion{};
    UINT iCopyrightLen{}, iProductVersionLen{};

    if (!VerQueryValue(data.data(), L"\\StringFileInfo\\040004b0\\LegalCopyright", (LPVOID*)&pvCopyright,
        &iCopyrightLen))
        std::abort();

    if (!VerQueryValue(data.data(), L"\\StringFileInfo\\040004b0\\ProductVersion", (LPVOID*)&pvProductVersion,
        &iProductVersionLen))
        std::abort();

    wcsncpy_s(szCopyright, uCopyrightLen, pvCopyright, iCopyrightLen);
    wcsncpy_s(szProductVersion, uProductVersionLen, pvProductVersion, iProductVersionLen);
}

INT_PTR CALLBACK about(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG: {
        WCHAR szCopyright[MAX_LOADSTRING]{};
        WCHAR szVersion[MAX_LOADSTRING]{};

        getProductAndVersion(szCopyright, ARRAYSIZE(szCopyright), szVersion, ARRAYSIZE(szVersion));

        SetWindowText(hDlg, g_szTitle);
        SetWindowText(GetDlgItem(hDlg, IDC_STATIC_COPYRIGHT), szCopyright);
        SetWindowText(GetDlgItem(hDlg, IDC_STATIC_VERSION), szVersion);

        return (INT_PTR)TRUE;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
