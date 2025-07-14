#include <windows.h>

// Convenience typedefs to define the arg types and such in one spot
typedef void(__stdcall* VbaCallback)(int);
typedef void(__stdcall* StringCallback)(const char*);
typedef void(__cdecl* RealRegister)(void* state, void (*cb)(int));
typedef void(__cdecl* RegisterMsgFunc)(void* state, void (*cb)(const char*));
typedef int(__cdecl* EnergyPlusFunc)(void* state, int argc, const char* argv[]);

// The callback instances need to be held somewhere -- here
static VbaCallback g_vbaCallback = NULL;
static StringCallback g_stringCallback = NULL;

// EnergyPlus ultimately actually calls these __cdecl stack functions
void __cdecl thunk_callback(int progress) {
    if (g_vbaCallback) {
        g_vbaCallback(progress);
    }
}
void __cdecl thunk_message_callback(const char* msg) {
    if (g_stringCallback) {
        g_stringCallback(msg);
    }
}

// These are the functions that VBA will call.  
// The arguments, including any VBA callbacks that are passed into here are __stdcall type, which E+ cannot handle.
// This layer insulates both sides from each other by using __cdecl args and registering a __cdecl callback as the actual callback from E+.
// Such fun.
__declspec(dllexport) void __stdcall trampolineRegisterProgressCallback(const char* apiPath, void* state, VbaCallback cb) {
    g_vbaCallback = cb;
    HMODULE hRealDLL = LoadLibraryA(apiPath);
    RealRegister realRegister = (RealRegister)GetProcAddress(hRealDLL, "registerProgressCallback");
    realRegister(state, thunk_callback);
}
__declspec(dllexport) void __stdcall trampolineRegisterStdOutCallback(const char* apiPath, void* state, StringCallback cb) {
    g_stringCallback = cb;
    HMODULE hDLL = LoadLibraryA(apiPath);
    RegisterMsgFunc realRegister = (RegisterMsgFunc)GetProcAddress(hDLL, "registerStdOutCallback");
    realRegister(state, thunk_message_callback);
}
__declspec(dllexport) int __stdcall trampolineEnergyPlus(const char* apiPath, void* state, int argc, const char** argv) {
    HMODULE hDLL = LoadLibraryA(apiPath);
    EnergyPlusFunc fn = (EnergyPlusFunc)GetProcAddress(hDLL, "energyplus");
    return fn(state, argc, argv);
}
