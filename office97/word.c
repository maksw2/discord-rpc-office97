#include <Windows.h>
#include "richpresence.h"

extern HWND g_hWnd;
extern WNDPROC g_OriginalWndProc;
extern BOOL g_IsInitialized;

LRESULT CALLBACK SubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
extern void UpdateTitle();
extern void Discord_Shutdown();
extern void discordInit(int app);

// WLL Auto-open function (called when WLL is loaded)
__declspec(dllexport) int __stdcall wdAutoOpen(void) {
    // Find Word's main window
    g_hWnd = FindWindowA("OpusApp", NULL);
    if (!g_hWnd) {
        MessageBoxA(NULL, "Could not find Word window", MESSAGER, MB_OK | MB_ICONERROR);
        return 0;
    }

    // Initialize Discord
    if (!g_IsInitialized) {
        discordInit(APP_WORD);
        g_IsInitialized = TRUE;

        // Subclass Word window to monitor title changes
        g_OriginalWndProc = (WNDPROC)SetWindowLongPtr(g_hWnd, GWLP_WNDPROC, (LONG_PTR)SubclassProc);

        UpdateTitle();

        MessageBoxA(g_hWnd, "Discord RPC initialized!\nYour Word activity will now show on Discord.",
            MESSAGER, MB_OK | MB_ICONINFORMATION);
    }

    return 1;
}

// WLL Auto-close function (called when WLL is unloaded)
__declspec(dllexport) int __stdcall wdAutoClose(void) {
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
