// main.cpp
#include <windows.h>
#include <string>
#include "ExcelCom.h"

#define HOTKEY_ID          1
#define WM_TRAYICON        (WM_USER + 1)
#define ID_TRAY_EXIT        1001
#define ID_TRAY_ABOUT       1002

ExcelCom g_excelApp;
NOTIFYICONDATA g_nid = { 0 };

// 添加系统托盘图标
void AddTrayIcon(HWND hwnd)
{
    g_nid.cbSize = sizeof(NOTIFYICONDATA);
    g_nid.hWnd = hwnd;
    g_nid.uID = 100;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wcscpy_s(g_nid.szTip, L"Excel Sheet Switcher - Ctrl+Alt+L");
    Shell_NotifyIcon(NIM_ADD, &g_nid);
}

void RemoveTrayIcon()
{
    Shell_NotifyIcon(NIM_DELETE, &g_nid);
}

void ShowContextMenu(HWND hwnd)
{
    HMENU hMenu = CreatePopupMenu();
    AppendMenu(hMenu, MF_STRING, ID_TRAY_ABOUT, L"Excel Sheet Switcher");
    AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hMenu, MF_STRING, ID_TRAY_EXIT, L"退出");

    POINT pt;
    GetCursorPos(&pt);
    SetForegroundWindow(hwnd);
    TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hwnd, NULL);
    DestroyMenu(hMenu);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CREATE:
    {
        // 尝试连接 Excel
        if (!g_excelApp.ConnectToRunningExcel())
        {
            MessageBox(hwnd, L"无法连接到正在运行的 Excel 实例！\n\n请先打开 Excel 再运行本程序。",
                       L"连接失败", MB_ICONWARNING);
            PostQuitMessage(1);
            return -1;
        }
        g_excelApp.RecordCurrentSheet();

        // 注册全局热键 Ctrl+Alt+L
        if (!RegisterHotKey(hwnd, HOTKEY_ID, MOD_CONTROL | MOD_ALT, 'L'))
        {
            MessageBox(hwnd, L"热键注册失败！\n\n请检查 Ctrl+Alt+L 是否被其他程序（如显卡驱动、输入法）占用。",
                       L"注册失败", MB_ICONERROR);
            PostQuitMessage(2);
            return -1;
        }

        AddTrayIcon(hwnd);
        return 0;
    }

    case WM_HOTKEY:
    {
        if (wParam == HOTKEY_ID)
        {
            g_excelApp.SwitchToPreviousSheet();
        }
        return 0;
    }

    case WM_TRAYICON:
    {
        if (lParam == WM_RBUTTONUP)
        {
            ShowContextMenu(hwnd);
        }
        return 0;
    }

    case WM_COMMAND:
    {
        if (LOWORD(wParam) == ID_TRAY_EXIT)
        {
            DestroyWindow(hwnd);
        }
        return 0;
    }

    case WM_DESTROY:
    {
        UnregisterHotKey(hwnd, HOTKEY_ID);
        RemoveTrayIcon();
        PostQuitMessage(0);
        return 0;
    }
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int)
{
    // 注册窗口类
    WNDCLASS wc = { 0 };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"ExcelSheetSwitcherClass";
    RegisterClass(&wc);

    // 创建隐藏窗口（仅用于接收消息）
    HWND hwnd = CreateWindow(wc.lpszClassName, L"Excel Sheet Switcher",
                             0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);
    if (!hwnd) return 3;

    // 消息循环
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}
