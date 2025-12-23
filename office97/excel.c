#include <Windows.h>
#include "richpresence.h"

extern HWND g_hWnd;
extern WNDPROC g_OriginalWndProc;
extern BOOL g_IsInitialized;

LRESULT CALLBACK SubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
extern void UpdateTitle();
extern void Discord_Shutdown();
extern void discordInit(int app);

// XLL Auto-open function (called when XLL is loaded)
__declspec(dllexport) int __stdcall xlAutoOpen(void) {
    // Find Excel's main window
    g_hWnd = FindWindowA("XLMAIN", NULL);
    if (!g_hWnd) {
        MessageBoxA(NULL, "Could not find Excel window", MESSAGER, MB_OK | MB_ICONERROR);
        return 0;
    }

    // Initialize Discord
    if (!g_IsInitialized) {
        discordInit(APP_EXCEL);
        g_IsInitialized = TRUE;

        // Subclass Excel window to monitor title changes
        g_OriginalWndProc = (WNDPROC)SetWindowLongPtr(g_hWnd, GWLP_WNDPROC, (LONG_PTR)SubclassProc);

        UpdateTitle();

        MessageBoxA(g_hWnd, "Discord RPC initialized!\nYour Excel activity will now show on Discord.",
            MESSAGER, MB_OK | MB_ICONINFORMATION);
    }

    return 1;
}

// XLL Auto-close function (called when XLL is unloaded)
__declspec(dllexport) int __stdcall xlAutoClose(void) {
    if (g_IsInitialized) {
        // Restore original window procedure
        if (g_OriginalWndProc && g_hWnd) {
            SetWindowLongPtr(g_hWnd, GWLP_WNDPROC, (LONG_PTR)g_OriginalWndProc);
        }

        Discord_Shutdown();
        g_IsInitialized = FALSE;
    }
    return 1;
}
