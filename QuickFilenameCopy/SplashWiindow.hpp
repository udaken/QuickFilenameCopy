#pragma once
#include <string>
#include <windows.h>
#include <wil/resource.h>


namespace
{
    using namespace std::string_literals;
    using namespace std::string_view_literals;

    class SplashWiindow
    {
        constexpr POINT centerOfRect(RECT rect)
        {
            return { rect.left + (rect.right - rect.left) / 2, rect.top + (rect.bottom - rect.top) / 2 };
        }
        static constexpr auto szWindowClassSplash = L"{D62A344C-F05F-44CD-BDA8-37038EF1E191}";

        static SplashWiindow* getSelf(HWND hWnd)
        {
            return (SplashWiindow*)GetWindowLongPtr(hWnd, GWLP_USERDATA);
        }

        void onPaint(wil::unique_hdc_paint hdc) noexcept
        {
            RECT rect{};
            GetClientRect(m_hWnd, &rect);

            FillRect(hdc.get(), &rect, (HBRUSH)GetStockObject(GRAY_BRUSH));
            SetBkMode(hdc.get(), TRANSPARENT);

            COLORREF bk_color = GetTextColor(hdc.get());
            auto textColor{ wil::scope_exit([&] { (void)SetTextColor(hdc.get(), bk_color); }) };
            SetTextColor(hdc.get(), RGB(255, 255, 255));

            auto hFontOld{ SelectObject(hdc.get(), m_hfont.get()) };
            auto font{ wil::scope_exit([&] { (void)SelectObject(hdc.get(), hFontOld); }) };
            DrawText(hdc.get(), L"Copied!", -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        }

        HWND m_hWnd{};
        UINT_PTR m_nRedrawTimerId{};
        UINT_PTR m_nCloseTimerId{};
        wil::unique_hfont m_hfont;

        static LRESULT CALLBACK wndProcSplash(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept
        {
            my::DbgPrint(__FUNCTIONW__ L"(hWnd={:p}, {})\n", (void*)hWnd, DebugPrintWndProc(message, wParam, lParam));

            switch (message)
            {
            case WM_NCCREATE:
                SetWindowLongPtr(hWnd, GWLP_USERDATA, LONG_PTR(((LPCREATESTRUCT)lParam)->lpCreateParams));
                getSelf(hWnd)->m_hWnd = hWnd;
                return DefWindowProc(hWnd, message, wParam, lParam);
            case WM_CREATE:
                getSelf(hWnd)->m_nRedrawTimerId = SetTimer(hWnd, 1, USER_TIMER_MINIMUM, nullptr);
                getSelf(hWnd)->m_nCloseTimerId = SetTimer(hWnd, 2, 500, nullptr);
                {
                    auto hFont = CreateFont(40, 20, 0,
                        FW_BOLD, FALSE, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, nullptr);
                    getSelf(hWnd)->m_hfont.reset(hFont);
                }

                return DefWindowProcW(hWnd, message, wParam, lParam);
            case WM_LBUTTONDOWN:
                DestroyWindow(hWnd);
                break;
            case WM_PAINT:
                getSelf(hWnd)->onPaint(wil::BeginPaint(hWnd));
                break;
            case WM_ERASEBKGND:
                return TRUE;
            case WM_TIMER:
                if (wParam == getSelf(hWnd)->m_nCloseTimerId)
                {
                    DestroyWindow(hWnd);
                }
                break;
            case WM_DESTROY:
                KillTimer(hWnd, getSelf(hWnd)->m_nRedrawTimerId);
                KillTimer(hWnd, getSelf(hWnd)->m_nCloseTimerId);
                return DefWindowProc(hWnd, message, wParam, lParam);
            case WM_NCDESTROY:
                getSelf(hWnd)->m_hWnd = nullptr;
                return DefWindowProc(hWnd, message, wParam, lParam);
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
            return 0;
        }

    public:

        [[nodiscard]]
        static WNDCLASSEXW wndClass() noexcept
        {
            
            WNDCLASSEXW wndClassSplash{
                DESIGNATED_INIT(.cbSize =) sizeof(wndClassSplash),
                DESIGNATED_INIT(.style =) CS_HREDRAW | CS_VREDRAW,
                DESIGNATED_INIT(.lpfnWndProc =) &wndProcSplash,
                DESIGNATED_INIT(.cbClsExtra =) 0,
                DESIGNATED_INIT(.cbWndExtra =) 0,
                DESIGNATED_INIT(.hInstance =) getHinstance(),
                DESIGNATED_INIT(.hIcon =) nullptr,
                DESIGNATED_INIT(.hCursor =) nullptr,
                DESIGNATED_INIT(.hbrBackground =) (HBRUSH)(COLOR_3DSHADOW),
                DESIGNATED_INIT(.lpszMenuName =) nullptr,
                DESIGNATED_INIT(.lpszClassName =) szWindowClassSplash,
            };
            return wndClassSplash;
        }

    public:
        ~SplashWiindow() noexcept
        {
            if (m_hWnd != nullptr)
            {
                std::terminate();
            }
        }

        void show(HINSTANCE hInstance, std::wstring_view g_szTitle)
        {
            TRACE();

            if (m_hWnd != nullptr)
                close();

            POINT center{ CW_USEDEFAULT , CW_USEDEFAULT };
            {
                POINT pos{};
                if (GetCursorPos(&pos)) {
                    auto hMonitor = MonitorFromPoint(pos, MONITOR_DEFAULTTONEAREST);
                    MONITORINFO mi{
                        DESIGNATED_INIT(.cbSize =) sizeof(mi),
                    };

                    if (GetMonitorInfo(hMonitor, &mi)) {
                        center = centerOfRect(mi.rcWork);
                    }
                }
            }

            const int width = 400;
            const int height = 400;

            HWND hWnd = CreateWindowEx(
                WS_EX_NOACTIVATE | WS_EX_COMPOSITED | WS_EX_LAYERED | WS_EX_TOPMOST,
                szWindowClassSplash, g_szTitle.data(), WS_POPUP | WS_VISIBLE,
                center.x - width / 2, center.y - height / 2,
                width, height, nullptr, nullptr, hInstance, this);
            THROW_LAST_ERROR_IF_NULL(hWnd);
        }

        void close()
        {
            if (m_hWnd != nullptr)
                DestroyWindow(m_hWnd);
        }
    };
}