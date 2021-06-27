#include "framework.h"
#include "resource.h"

#include <unordered_map>
#include <wil/com.h>
#include <wil/resource.h>
#include <wil/result.h>

#pragma comment(lib, "Version.lib")
#pragma comment(lib, "Shlwapi.lib")

#include "DebugPrintWndProc.hpp"

#pragma comment(linker, "/manifestdependency:\"type='win32' \
    name='Microsoft.Windows.Common-Controls' \
    version='6.0.0.0' \
    processorArchitecture='*' \
    publicKeyToken='6595b64144ccf1df' \
    language='*'\"")

#define WM_NOTIFYICON (WM_USER + 100)
constexpr int TIMER_ID_ADDTRAYICON = 100;

#define CATCH_SHOW_MSGBOX(hWnd)                                                     \
    catch (const wil::ResultException &e)                                           \
    {                                                                               \
        wchar_t message[2048]{};                                                    \
        wil::GetFailureLogString(message, ARRAYSIZE(message), e.GetFailureInfo());  \
        ::MessageBox(hWnd, message, g_szTitle.c_str(), MB_ICONERROR);                         \
    }

#define TRACE() my::DbgPrint("{}", __FUNCTION__ "\n")
//#define TRACE() 
#define DBGPRINTLN(fmt, ...) my::DbgPrint(__FILE__ "(" _STRINGIZE(__LINE__) "): " fmt "\n", __VA_ARGS__)

#if __cpp_designated_initializers
#define DESIGNATED_INIT(designator) designator
#else
#define DESIGNATED_INIT(designator) 
#endif

using namespace std::string_literals;
using namespace std::string_view_literals;

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

    inline std::wstring loadString(HINSTANCE hInstance, UINT uID)
    {
        LPCVOID dummy{};
        size_t count = static_cast<size_t>(LoadStringW(hInstance, uID, (PWSTR)&dummy, 0));
        std::wstring buffer(count, L'\0');
        if (::LoadStringW(hInstance, uID, buffer.data(), static_cast<int>(count) + 1) > 0)
        {
            return buffer;
        }
        else
        {
            return {};
        }
    }

} // namespace my

static HINSTANCE getHinstance() noexcept
{
    extern HINSTANCE g_hInst;

    if (g_hInst == nullptr)
        std::terminate();

    return g_hInst;
}

#include "SplashWiindow.hpp"

HHOOK g_hook;
std::wstring g_szTitle;
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
    using result = typename err_policy::result;

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
    if (ss.empty()) {
        return;
    }

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
        hGlobal.release();
    }
}

SplashWiindow g_splashWindow;

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
    g_splashWindow.show(getHinstance(), (g_szTitle + L" Splash"s).c_str());
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

            my::DbgPrint(L"[{}] name={}, fullname={}, hwnd={}\n"sv, i, name.get(), fullname.get(), hWndShell);
        }

        if (reinterpret_cast<HWND>(hWndShell) == hWndTarget)
        {
            auto psp = pWBA.query<IServiceProvider>();

            wil::com_ptr_t<IShellBrowser> psb;
            THROW_IF_FAILED(psp->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&psb)));

            wil::com_ptr_t<IShellView> psv;
            THROW_IF_FAILED(psb->QueryActiveShellView(&psv));

            auto pfv2 = psv.query<IFolderView2>();

            { // debug

                UINT viewMode{};
                THROW_IF_FAILED(pfv2->GetCurrentViewMode(&viewMode));

                int itemCount{};
                THROW_IF_FAILED(pfv2->ItemCount(SVGIO_BACKGROUND, &itemCount));

                my::DbgPrint(L"[{}] viewMode={}, itemCount={}\n"sv, i, viewMode, itemCount);
            }

            copySelectedItems(pfv2);
            my::DbgPrint(L"[{}] Copied!\n"sv, i);
            break;
        }
    }
}

LRESULT CALLBACK lowLevelKeyboardProc(int code, WPARAM wParam, LPARAM lParam) noexcept
{
    TRACE();

    if (code < HC_ACTION)
        return CallNextHookEx(nullptr, code, wParam, lParam);

    auto pKbdll = reinterpret_cast<const KBDLLHOOKSTRUCT*>(lParam);

    DBGPRINTLN("flags:{:x}, vkCode:{:x}"sv, pKbdll->flags, pKbdll->vkCode);

    if ((pKbdll->flags & (LLKHF_LOWER_IL_INJECTED | LLKHF_UP)) == 0 && pKbdll->vkCode == 'C')
    {
        auto ctrlKey = GetAsyncKeyState(VK_CONTROL) & 0x8000;
        auto shiftKey = GetAsyncKeyState(VK_SHIFT) & 0x8000;
        DBGPRINTLN("ctrlKey:{}, shiftKey:{}"sv, ctrlKey, shiftKey);

        if (ctrlKey != 0 && shiftKey != 0)
        {
            auto hWnd{ GetForegroundWindow() };
            if (hWnd == nullptr) {
                hWnd = GetForegroundWindow();
            }

            if (hWnd != nullptr)
            {
                WCHAR className[512]{};
                RealGetWindowClass(hWnd, className, ARRAYSIZE(className));

                DBGPRINTLN(L"className:{}", className);
                if (lstrcmp(className, targetClassName) == 0)
                    try {
                    traverseShellWindows(g_pSHWinds, hWnd);
                    return 1;
                }
                catch (...)
                {
                    DBGPRINTLN("catch");
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

bool addNotifyIcon(HWND hWnd, unsigned int uID)
{
    TRACE();
    NOTIFYICONDATA nid{ sizeof(nid) };

    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.hWnd = hWnd;
    nid.uID = uID;
    nid.uCallbackMessage = WM_NOTIFYICON;
    nid.hIcon = LoadIcon(getHinstance(), MAKEINTRESOURCE(IDI_SMALL));
    g_szTitle.copy(nid.szTip, std::size(nid.szTip));

    BOOL success = Shell_NotifyIcon(NIM_ADD, &nid);
    THROW_LAST_ERROR_IF(success == FALSE && GetLastError() != ERROR_TIMEOUT);

    return success != FALSE;
}

void tryAddNotifyIcon(HWND hWnd, unsigned int uID)
{
    if (!addNotifyIcon(hWnd, uID))
    {
        THROW_LAST_ERROR_IF(SetTimer(hWnd, TIMER_ID_ADDTRAYICON, 1000, nullptr) == 0);
    }
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

    WNDCLASSEXW wndClass{
        DESIGNATED_INIT(.cbSize =) sizeof(wndClass),
        DESIGNATED_INIT(.style =) CS_HREDRAW | CS_VREDRAW,
        DESIGNATED_INIT(.lpfnWndProc =) &wndProc,
        DESIGNATED_INIT(.cbClsExtra = ) 0,
        DESIGNATED_INIT(.cbWndExtra = ) 0,
        DESIGNATED_INIT(.hInstance =) hInstance,
        DESIGNATED_INIT(.hIcon =) LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL)),
        DESIGNATED_INIT(.hCursor = ) nullptr,
        DESIGNATED_INIT(.hbrBackground = ) nullptr,
        DESIGNATED_INIT(.lpszMenuName = ) nullptr,
        DESIGNATED_INIT(.lpszClassName =) szWindowClass,
        DESIGNATED_INIT(.hIconSm =) LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL)),
    };

    WNDCLASSEXW wndClassSplash{ SplashWiindow::wndClass() };

    return RegisterClassExW(&wndClass) && RegisterClassExW(&wndClassSplash);
}

HINSTANCE g_hInst;

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine,
                      _In_ int nCmdShow)
{
    SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_SYSTEM32);

    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);
    UNREFERENCED_PARAMETER(nCmdShow);

    g_hInst = hInstance;

    g_szTitle = my::loadString(hInstance, IDS_APP_TITLE);

    g_szTitle.append(
#if _M_ARM64 
             L" (ARM64bit)"
#elif _M_X64
             L" (64bit)"
#elif _M_IX86 
             L" (32bit)"
#endif
    );

    wil::unique_mutex m{ CreateMutex(nullptr, FALSE, g_szTitle.c_str()) };
    if (GetLastError() == ERROR_ALREADY_EXISTS)
    {
        ::MessageBox(nullptr, L"Another instance is already running.", g_szTitle.c_str(), MB_ICONEXCLAMATION);
        return 0;
    }
    auto hr{ CoInitializeEx(nullptr, COINIT_MULTITHREADED | COINIT_DISABLE_OLE1DDE) };
    THROW_IF_FAILED(hr);
    auto initialized = wil::scope_exit([] { CoUninitialize(); });

    g_pSHWinds = wil::CoCreateInstance<ShellWindows, IShellWindows>(CLSCTX_ALL);

    installHook();
    auto hook = wil::scope_exit([] {
        uninstallHook();
        g_pSHWinds = nullptr;
    });

    if (registerMyClass(hInstance) == 0)
    {
        THROW_LAST_ERROR();
    }

    HWND hWnd = CreateWindowW(szWindowClass, g_szTitle.c_str(), 0, CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr,
                              hInstance, nullptr);

    if (hWnd == nullptr)
    {
        return FALSE;
    }

    tryAddNotifyIcon(hWnd, NOTIFY_UID);

    MSG msg{};
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    deleteNotifyIcon(hWnd, NOTIFY_UID);

    return (int)msg.wParam;
}

void registerToShortcut(HWND hWnd)
try
{
    WCHAR targetPath[MAX_PATH]{};
    THROW_IF_WIN32_BOOL_FALSE(SHGetSpecialFolderPath(hWnd, targetPath, CSIDL_STARTUP, FALSE));

    //THROW_IF_FAILED(StringCchCat(targetPath, ARRAYSIZE(targetPath), ((L"\\"sv + g_szTitle).append(L".lnk"s)).c_str()));

    auto pShellLink{ wil::CoCreateInstance<IShellLink>(CLSID_ShellLink) };
    auto pPersistFile{ pShellLink.query<IPersistFile>() };

    THROW_IF_FAILED(pShellLink->SetPath(my::GetModuleFileName().c_str()));
    THROW_IF_FAILED(pPersistFile->Save((targetPath + L"\\"s + g_szTitle + L".lnk"s).c_str(), TRUE));
}
CATCH_SHOW_MSGBOX(hWnd)

LRESULT CALLBACK wndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    static UINT s_uTaskbarRestart;
    static HMENU s_menu = nullptr;

    switch (message)
    {
    case WM_CREATE:
        g_hwnd = hWnd;
        s_uTaskbarRestart = RegisterWindowMessage(L"TaskbarCreated");
        s_menu = LoadMenu(getHinstance(), MAKEINTRESOURCE(IDR_MENU1));
        return DefWindowProc(hWnd, message, wParam, lParam);
    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case ID_ROOT_ABOUT:
            DialogBox(getHinstance(), MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, &about);
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
            TrackPopupMenu(GetSubMenu(s_menu, 0), TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, nullptr);
        }
        break;
        default:
            ;
        }
        return 0;
    case WM_TIMER:
        switch (wParam)
        {
        case TIMER_ID_ADDTRAYICON:
            KillTimer(hWnd, TIMER_ID_ADDTRAYICON);
            tryAddNotifyIcon(hWnd, NOTIFY_UID);
            break;
        default:
            break;
        }
        return DefWindowProc(hWnd, message, wParam, lParam);
    default:
        if (message == s_uTaskbarRestart)
        {
            tryAddNotifyIcon(hWnd, NOTIFY_UID);
        }
        else
            return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

auto getProductAndVersion()
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

    if (!VerQueryValue(data.data(), LR"(\StringFileInfo\040004b0\LegalCopyright)", (LPVOID*)&pvCopyright,
        &iCopyrightLen))
        std::abort();

    if (!VerQueryValue(data.data(), LR"(\StringFileInfo\040004b0\ProductVersion)", (LPVOID*)&pvProductVersion,
        &iProductVersionLen))
        std::abort();

    struct {
        std::wstring copyright;
        std::wstring productVersion;
    } result = {
        std::wstring { pvCopyright ,iCopyrightLen },
        std::wstring { pvProductVersion ,iProductVersionLen },
    };

    return result;
}

INT_PTR CALLBACK about(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) noexcept
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG: {
        auto [copyright, version] = getProductAndVersion();

        THROW_IF_WIN32_BOOL_FALSE(SetWindowText(hDlg, g_szTitle.c_str()));
        THROW_IF_WIN32_BOOL_FALSE(SetDlgItemText(hDlg, IDC_STATIC_COPYRIGHT, copyright.data()));
        THROW_IF_WIN32_BOOL_FALSE(SetDlgItemText(hDlg, IDC_STATIC_VERSION, version.data()));

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
