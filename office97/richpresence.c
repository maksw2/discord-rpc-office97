#include <Windows.h>
#include <wchar.h>
#include <stdint.h>
#include <stdio.h>
#include <string.h>
#include <time.h>
#include <discord_rpc.h>
#include "richpresence.h"

static const char* EXCEL_APPLICATION_ID = "1452387313530441932";
static const char* WORD_APPLICATION_ID = "1452662950425919608";
const char* MESSAGER = "Office 97 Discord RPC";
static int64_t StartTime;
extern char title[512];
extern HWND g_hWnd;
extern BOOL g_IsInitialized;

void updateDiscordPresence() {
    DiscordRichPresence discordPresence = { 0 };
    discordPresence.details = title;
    discordPresence.startTimestamp = StartTime;
    discordPresence.endTimestamp = time(0) + 5 * 60;
    discordPresence.largeImageKey = "canary-large";
    discordPresence.smallImageKey = "ptb-small";
    discordPresence.instance = 0;
    Discord_UpdatePresence(&discordPresence);
}

static void handleDiscordReady(const DiscordUser* connectedUser) {
    char buffer[256];
    sprintf_s(buffer, sizeof(buffer), "Discord: connected to user %s#%s - %s",
        connectedUser->username,
        connectedUser->discriminator,
        connectedUser->userId);
    MessageBoxA(g_hWnd, buffer, MESSAGER, MB_OK | MB_ICONINFORMATION);
}

static void handleDiscordDisconnected(int errcode, const char* message) {
    char buffer[256];
    sprintf_s(buffer, sizeof(buffer), "Discord: disconnected (%d: %s)", errcode, message);
    MessageBoxA(g_hWnd, buffer, MESSAGER, MB_OK | MB_ICONINFORMATION);
}

static void handleDiscordError(int errcode, const char* message) {
    char buffer[256];
    sprintf_s(buffer, sizeof(buffer), "Discord: error (%d: %s)", errcode, message);
    MessageBoxA(g_hWnd, buffer, MESSAGER, MB_OK | MB_ICONERROR);
}

void discordInit(int app) {
    DiscordEventHandlers handlers = { 0 };
    handlers.ready = handleDiscordReady;
    handlers.disconnected = handleDiscordDisconnected;
    handlers.errored = handleDiscordError;

    if (app == APP_EXCEL)
        Discord_Initialize(EXCEL_APPLICATION_ID, &handlers, 1, NULL);
    else if (app == APP_WORD)
        Discord_Initialize(WORD_APPLICATION_ID, &handlers, 1, NULL);
    else {
        MessageBoxA(g_hWnd, "Wrong app id passed to discordInit(), terminating.", MESSAGER, MB_OK | MB_ICONERROR);
        abort();
    }

    StartTime = time(0);
}

// Timer callback for Discord callbacks
static VOID CALLBACK TimerProc([[maybe_unused]] HWND hwnd, [[maybe_unused]] UINT uMsg, [[maybe_unused]] UINT_PTR idEvent, [[maybe_unused]] DWORD dwTime) {
    Discord_RunCallbacks();
}

// DLL entry point
BOOL APIENTRY DllMain([[maybe_unused]] HMODULE hModule, DWORD ul_reason_for_call, [[maybe_unused]] LPVOID lpReserved) {
    switch (ul_reason_for_call) {
    case DLL_PROCESS_ATTACH:
        // Set up a timer for Discord callbacks
        SetTimer(NULL, 0, 1000, TimerProc); // Call every second
        break;
    case DLL_PROCESS_DETACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
        break;
    }
    return TRUE;
}
