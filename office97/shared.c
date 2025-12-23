#include <Windows.h>

HWND g_hWnd = NULL;
WNDPROC g_OriginalWndProc = NULL;
BOOL g_IsInitialized = FALSE;
char title[512];

extern void updateDiscordPresence();

// Get the active workbook/document name
void UpdateTitle() {
    if (g_hWnd) {
        wchar_t windowTitle[512];
        GetWindowTextW(g_hWnd, windowTitle, sizeof(windowTitle) / sizeof(wchar_t));

        // Find the first occurrence of the hyphen
        wchar_t* hyphen = wcschr(windowTitle, L'-');

        if (hyphen) {
            // Move pointer past the hyphen
            hyphen++;

            // Skip any leading spaces that follow the hyphen (e.g., " -  filename.xls")
            while (*hyphen == L' ') {
                hyphen++;
            }

            // Convert the remaining string starting from 'hyphen' to UTF-8
            WideCharToMultiByte(CP_UTF8, 0, hyphen, -1, title, sizeof(title), NULL, NULL);
        }
        else {
            // If no hyphen is found, use the whole title as a fallback
            WideCharToMultiByte(CP_UTF8, 0, windowTitle, -1, title, sizeof(title), NULL, NULL);
        }

        updateDiscordPresence();
    }
}

// Subclass Excel/Word's window to detect workbook/document changes
LRESULT CALLBACK SubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    if (uMsg == WM_SETTEXT || uMsg == WM_ACTIVATE || uMsg == WM_ACTIVATEAPP) {
        UpdateTitle();
    }
    return CallWindowProc(g_OriginalWndProc, hwnd, uMsg, wParam, lParam);
}
