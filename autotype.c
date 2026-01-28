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

TCHAR g_szAppName[] = _T("AutoType");
TCHAR g_szStatic[] = _T("STATIC");
TCHAR g_szEdit[] = _T("EDIT");
TCHAR g_szLabelSpeed[] = _T("Interval (ms):");
TCHAR g_szHelpText[] = _T("1. Copy text\r\n2. Ctrl+Alt+V to type\r\n3. Esc to interrupt");

int g_nIntervalMs = 15;
BOOL g_bIsEnabled = TRUE;
volatile BOOL g_bIsTyping = FALSE;
NOTIFYICONDATA g_nidTray = {0};
HWND g_hwndMain, g_hwndConfig, g_hwndHelp, g_hwndUpdown;
HANDLE g_hMutex;

void SendKey(wchar_t wch) {
    INPUT in[2] = {0};
    SHORT ks;
    ks = VkKeyScanW(wch);
    if (ks != -1) {
        BYTE bSh = HIBYTE(ks);
        WORD wSc = (WORD)MapVirtualKey(LOBYTE(ks), 0);
        if (bSh & 1) {
            INPUT si = {0};
            si.type = INPUT_KEYBOARD;
            si.ki.wScan = (WORD)MapVirtualKey(VK_SHIFT, 0);
            si.ki.dwFlags = KEYEVENTF_SCANCODE;
            SendInput(1, &si, sizeof(INPUT));
        }
        in[0].type = in[1].type = INPUT_KEYBOARD;
        in[0].ki.wScan = in[1].ki.wScan = wSc;
        in[0].ki.dwFlags = KEYEVENTF_SCANCODE;
        in[1].ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;
        SendInput(2, in, sizeof(INPUT));
        if (bSh & 1) {
            INPUT si = {0};
            si.type = INPUT_KEYBOARD;
            si.ki.wScan = (WORD)MapVirtualKey(VK_SHIFT, 0);
            si.ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;
            SendInput(1, &si, sizeof(INPUT));
        }
    } else {
        in[0].type = in[1].type = INPUT_KEYBOARD;
        in[0].ki.wScan = in[1].ki.wScan = wch;
        in[0].ki.dwFlags = KEYEVENTF_UNICODE;
        in[1].ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
        SendInput(2, in, sizeof(INPUT));
    }
}

DWORD WINAPI TypeThreadProc(LPVOID lpParam) {
    wchar_t* p = (wchar_t*)lpParam;
    wchar_t* cur = p;
    while ((GetAsyncKeyState(VK_CONTROL) & 0x8000) || (GetAsyncKeyState(VK_MENU) & 0x8000)) Sleep(10);
    while (*cur && g_bIsTyping) {
        if (GetAsyncKeyState(VK_ESCAPE) & 0x8000) break;
        if (*cur != L'\r') {
            SendKey(*cur);
            if (g_nIntervalMs > 0) Sleep(g_nIntervalMs + (GetTickCount() & 7));
        }
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
                        size_t sz = (wcslen(pc) + 1) * sizeof(wchar_t);
                        wchar_t* pBuf = (wchar_t*)GlobalAlloc(GPTR, sz);
                        if (pBuf) {
                            HANDLE hThread;
                            CopyMemory(pBuf, pc, sz);
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

int WINAPI WinMain(HINSTANCE hI, HINSTANCE hP, LPSTR lpC, int nS) {
    WNDCLASS wc = {0};
    INITCOMMONCONTROLSEX ic;
    HWND he;
    MSG msg;

    g_hMutex = CreateMutex(NULL, TRUE, g_szAppName);
    if (GetLastError() == ERROR_ALREADY_EXISTS) return 0;

    wc.lpfnWndProc = WndProc;
    wc.hInstance = hI;
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = g_szAppName;
    RegisterClass(&wc);

    g_hwndMain = CreateWindowEx(0, g_szAppName, g_szAppName, 0, 0, 0, 0, 0, NULL, NULL, hI, NULL);
    g_hwndConfig = CreateWindowEx(WS_EX_TOPMOST | WS_EX_TOOLWINDOW, g_szAppName, g_szAppName, WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, 0, 0, 0, 0, NULL, NULL, hI, NULL);
    g_hwndHelp = CreateWindowEx(WS_EX_TOPMOST | WS_EX_TOOLWINDOW, g_szAppName, g_szAppName, WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, 0, 0, 0, 0, NULL, NULL, hI, NULL);

    ic.dwSize = sizeof(ic);
    ic.dwICC = ICC_UPDOWN_CLASS;
    InitCommonControlsEx(&ic);

    CreateWindow(g_szStatic, g_szLabelSpeed, WS_VISIBLE | WS_CHILD, 15, 22, 90, 20, g_hwndConfig, NULL, hI, NULL);
    he = CreateWindowEx(WS_EX_CLIENTEDGE, g_szEdit, NULL, WS_VISIBLE | WS_CHILD | ES_NUMBER | ES_RIGHT, 110, 20, 55, 25, g_hwndConfig, NULL, hI, NULL);
    g_hwndUpdown = CreateWindowEx(0, UPDOWN_CLASS, NULL, WS_VISIBLE | WS_CHILD | UDS_SETBUDDYINT | UDS_ALIGNRIGHT | UDS_ARROWKEYS | UDS_NOTHOUSANDS, 0, 0, 0, 0, g_hwndConfig, NULL, hI, NULL);
    SendMessage(g_hwndUpdown, UDM_SETBUDDY, (WPARAM)he, 0);
    SendMessage(g_hwndUpdown, UDM_SETRANGE, 0, MAKELPARAM(9999, 0));
    SendMessage(g_hwndUpdown, UDM_SETPOS, 0, (LPARAM)g_nIntervalMs);

    CreateWindow(g_szStatic, g_szHelpText, WS_VISIBLE | WS_CHILD, 15, 15, 190, 60, g_hwndHelp, NULL, hI, NULL);

    RegisterHotKey(g_hwndMain, HK_ID, MOD_CONTROL | MOD_ALT, 0x56);
    
    g_nidTray.cbSize = sizeof(NOTIFYICONDATA);
    g_nidTray.hWnd = g_hwndMain;
    g_nidTray.uID = 1;
    g_nidTray.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nidTray.uCallbackMessage = WM_TRAY;
    g_nidTray.hIcon = ExtractIcon(hI, _T("pifmgr.dll"), 12);
    if (!g_nidTray.hIcon) g_nidTray.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    _tcscpy(g_nidTray.szTip, g_szAppName);
    Shell_NotifyIcon(NIM_ADD, &g_nidTray);

    while (GetMessage(&msg, NULL, 0, 0)) { TranslateMessage(&msg); DispatchMessage(&msg); }
    if (g_hMutex) { ReleaseMutex(g_hMutex); CloseHandle(g_hMutex); }
    return (int)msg.wParam;
}
