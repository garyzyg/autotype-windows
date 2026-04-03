#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shellapi.h>
#include <commctrl.h>
#include <tchar.h>

#ifndef KEYEVENTF_UNICODE
#define KEYEVENTF_UNICODE 0x0004
#endif

#define HK_ID 1001
#define WM_TRAY (WM_USER + 10)
#define ID_TOGGLE 2001
#define ID_CONFIG 2002
#define ID_HELP   2003
#define ID_EXIT   2004

#ifdef RtlMoveMemory
#undef RtlMoveMemory
WINBASEAPI VOID WINAPI RtlMoveMemory(PVOID Destination, const VOID* Source, SIZE_T Length);
#endif

#ifdef RtlZeroMemory
#undef RtlZeroMemory
WINBASEAPI VOID WINAPI RtlZeroMemory(PVOID Destination, SIZE_T Length);
#endif

int g_nIntervalMs = 15;
BOOL g_bIsEnabled = TRUE;
volatile BOOL g_bIsTyping = FALSE;
NOTIFYICONDATA g_nidTray;
HWND g_hwndMain, g_hwndConfig, g_hwndHelp, g_hwndUpdown;

void SendKey(wchar_t wch) {
    SHORT ks = VkKeyScanW(wch);
    UINT count = 0;
    UINT i = 0;
    INPUT in[4];
    RtlZeroMemory(in, sizeof(in));

    for (i = 0; i < ARRAYSIZE(in); i++) {
        in[i].type = INPUT_KEYBOARD;
    }

    if (ks != -1) {
        BYTE bSh = HIBYTE(ks);
        WORD wSc = (WORD)MapVirtualKey(LOBYTE(ks), 0);

        if (bSh & 1) {
            in[count].ki.wScan = (WORD)MapVirtualKey(VK_SHIFT, 0);
            in[count].ki.dwFlags = KEYEVENTF_SCANCODE;
            count++;
        }

        in[count].ki.wScan = wSc;
        in[count].ki.dwFlags = KEYEVENTF_SCANCODE;
        count++;

        in[count].ki.wScan = wSc;
        in[count].ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;
        count++;

        if (bSh & 1) {
            in[count].ki.wScan = (WORD)MapVirtualKey(VK_SHIFT, 0);
            in[count].ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;
            count++;
        }
    } else {
        in[0].ki.wScan = in[1].ki.wScan = (WORD)wch;
        in[0].ki.dwFlags = KEYEVENTF_UNICODE;
        in[1].ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
        count = 2;
    }

    if (count > 0) {
        SendInput(count, in, sizeof(INPUT));
    }
}

DWORD WINAPI TypeThreadProc(LPVOID lpParam) {
    wchar_t* p = (wchar_t*)lpParam;
    wchar_t* cur = p;
    while ((GetAsyncKeyState(VK_CONTROL) & 0x8000) || (GetAsyncKeyState(VK_MENU) & 0x8000)) Sleep(10);
    while (*cur && g_bIsTyping) {
        if (GetAsyncKeyState(VK_ESCAPE) & 0x8000) break;
        if (*cur == L'\r') {
            if (*(cur + 1) == L'\n') cur++;
            SendKey(L'\r');
        } else if (*cur == L'\n') {
            SendKey(L'\r');
        } else {
            SendKey(*cur);
        }
        if (g_nIntervalMs > 0) Sleep(g_nIntervalMs + (GetTickCount() & 7));
        cur++;
    }
    GlobalFree(p);
    g_bIsTyping = FALSE;
    return 0;
}

void ShowCentered(HWND h, int w, int h_c) {
    RECT r = {0, 0, w, h_c};
    int nw, nh, px, py;
    AdjustWindowRectEx(&r, (DWORD)GetWindowLongPtr(h, GWL_STYLE), FALSE, (DWORD)GetWindowLongPtr(h, GWL_EXSTYLE));
    nw = r.right - r.left; nh = r.bottom - r.top;
    px = (GetSystemMetrics(SM_CXSCREEN) - nw) / 2;
    py = (GetSystemMetrics(SM_CYSCREEN) - nh) / 2;
    SetWindowPos(h, HWND_TOPMOST, px, py, nw, nh, SWP_SHOWWINDOW);
    SetForegroundWindow(h);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT m, WPARAM w, LPARAM l) {
    switch (m) {
    case WM_TRAY:
        if (l == WM_RBUTTONUP) {
            HMENU hm = CreatePopupMenu();
            POINT pt;
            TCHAR buf[32];
            GetCursorPos(&pt);
            AppendMenu(hm, MF_STRING | (g_bIsEnabled ? MF_CHECKED : 0), ID_TOGGLE, _T("Enabled"));
            wsprintf(buf, _T("Interval: %d ms"), g_nIntervalMs);
            AppendMenu(hm, MF_STRING, ID_CONFIG, buf);
            AppendMenu(hm, MF_SEPARATOR, 0, 0);
            AppendMenu(hm, MF_STRING, ID_HELP, _T("Help"));
            AppendMenu(hm, MF_STRING, ID_EXIT, _T("Exit"));
            SetForegroundWindow(hWnd);
            TrackPopupMenu(hm, TPM_RIGHTBUTTON, pt.x, pt.y, 0, g_hwndMain, 0);
            DestroyMenu(hm);
        }
        break;
    case WM_HOTKEY:
        if (w == HK_ID && g_bIsEnabled && !g_bIsTyping) {
            if (OpenClipboard(NULL)) {
                HANDLE hd = GetClipboardData(CF_UNICODETEXT);
                if (hd) {
                    wchar_t* pc = (wchar_t*)GlobalLock(hd);
                    if (pc) {
                        size_t sz = (lstrlenW(pc) + 1) * sizeof(wchar_t);
                        wchar_t* pBuf = (wchar_t*)GlobalAlloc(GPTR, sz);
                        if (pBuf) {
                            HANDLE hThread;
                            RtlMoveMemory(pBuf, pc, sz);
                            g_bIsTyping = TRUE;
                            hThread = CreateThread(NULL, 0, TypeThreadProc, pBuf, 0, NULL);
                            if (hThread) CloseHandle(hThread);
                        }
                        GlobalUnlock(hd);
                    }
                }
                CloseClipboard();
            }
        }
        break;
    case WM_COMMAND:
        switch (LOWORD(w)) {
        case ID_TOGGLE: g_bIsEnabled = !g_bIsEnabled; break;
        case ID_CONFIG: ShowCentered(g_hwndConfig, 185, 70); break;
        case ID_HELP:   ShowCentered(g_hwndHelp, 220, 80); break;
        case ID_EXIT:   DestroyWindow(g_hwndMain); break;
        }
        break;
    case WM_CLOSE:
        if (hWnd == g_hwndConfig) g_nIntervalMs = (int)SendMessage(g_hwndUpdown, UDM_GETPOS, 0, 0) & 0xFFFF;
        ShowWindow(hWnd, SW_HIDE);
        return 0;
    case WM_DESTROY:
        if (hWnd == g_hwndMain) {
            Shell_NotifyIcon(NIM_DELETE, &g_nidTray);
            if (g_nidTray.hIcon) DestroyIcon(g_nidTray.hIcon);
            UnregisterHotKey(hWnd, HK_ID);
            PostQuitMessage(0);
        }
        break;
    default: return DefWindowProc(hWnd, m, w, l);
    }
    return 0;
}

void _start() {
    HINSTANCE hI = GetModuleHandle(NULL);
    static const TCHAR szAppName[] = _T("AutoType");
    static const TCHAR szStatic[] = _T("STATIC");
    INITCOMMONCONTROLSEX ic;
    HWND he;
    MSG msg;
    WNDCLASS wc;
    RtlZeroMemory(&wc, sizeof(wc));
    RtlZeroMemory(&g_nidTray, sizeof(g_nidTray));

    CreateMutex(NULL, TRUE, szAppName);
    if (GetLastError() == ERROR_ALREADY_EXISTS) ExitProcess(0);

    wc.lpfnWndProc = WndProc;
    wc.hInstance = hI;
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = szAppName;
    RegisterClass(&wc);

    g_hwndMain = CreateWindowEx(0, szAppName, szAppName, 0, 0, 0, 0, 0, NULL, NULL, hI, NULL);
    g_hwndConfig = CreateWindowEx(WS_EX_TOPMOST | WS_EX_TOOLWINDOW, szAppName, szAppName, WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, 0, 0, 0, 0, NULL, NULL, hI, NULL);
    g_hwndHelp = CreateWindowEx(WS_EX_TOPMOST | WS_EX_TOOLWINDOW, szAppName, szAppName, WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, 0, 0, 0, 0, NULL, NULL, hI, NULL);

    ic.dwSize = sizeof(ic);
    ic.dwICC = ICC_UPDOWN_CLASS;
    InitCommonControlsEx(&ic);

    CreateWindow(szStatic, _T("Interval (ms):"), WS_VISIBLE | WS_CHILD, 15, 22, 90, 20, g_hwndConfig, NULL, hI, NULL);
    he = CreateWindowEx(WS_EX_CLIENTEDGE, _T("EDIT"), NULL, WS_VISIBLE | WS_CHILD | ES_NUMBER | ES_RIGHT, 110, 20, 55, 25, g_hwndConfig, NULL, hI, NULL);
    g_hwndUpdown = CreateWindowEx(0, UPDOWN_CLASS, NULL, WS_VISIBLE | WS_CHILD | UDS_SETBUDDYINT | UDS_ALIGNRIGHT | UDS_ARROWKEYS | UDS_NOTHOUSANDS, 0, 0, 0, 0, g_hwndConfig, NULL, hI, NULL);
    SendMessage(g_hwndUpdown, UDM_SETBUDDY, (WPARAM)he, 0);
    SendMessage(g_hwndUpdown, UDM_SETRANGE, 0, MAKELPARAM(9999, 0));
    SendMessage(g_hwndUpdown, UDM_SETPOS, 0, (LPARAM)g_nIntervalMs);

    CreateWindow(szStatic, _T("1. Copy text\r\n2. Ctrl+Alt+V to type\r\n3. Esc to interrupt"), WS_VISIBLE | WS_CHILD, 15, 15, 190, 60, g_hwndHelp, NULL, hI, NULL);

    RegisterHotKey(g_hwndMain, HK_ID, MOD_CONTROL | MOD_ALT, 0x56);
    
    g_nidTray.cbSize = sizeof(NOTIFYICONDATA);
    g_nidTray.hWnd = g_hwndMain;
    g_nidTray.uID = 1;
    g_nidTray.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nidTray.uCallbackMessage = WM_TRAY;
    g_nidTray.hIcon = ExtractIcon(hI, _T("pifmgr.dll"), 12);
    if (!g_nidTray.hIcon) g_nidTray.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    lstrcpyn(g_nidTray.szTip, szAppName, ARRAYSIZE(g_nidTray.szTip));
    Shell_NotifyIcon(NIM_ADD, &g_nidTray);

    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    ExitProcess(0);
}
