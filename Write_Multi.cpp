// Write_Multi.cpp: ファイルを重複保存できるEDCB/TVTestの出力プラグイン
// Public domain
#define WIN32_LEAN_AND_MEAN
#include <Windows.h>
#include <shellapi.h>

// 同時にCreateCtrl()できる数
#define MAX_CONTEXT 16
// 同時にDuplicateSave()できる数
#define MAX_DUP 64
// 扱えるパスの長さ
#define MY_MAX_PATH 512

#define DLL_EXPORT extern "C" __declspec(dllexport)

class CBlockLock
{
public:
    CBlockLock(CRITICAL_SECTION *lock) : m_lock(lock) { EnterCriticalSection(m_lock); }
    ~CBlockLock() { LeaveCriticalSection(m_lock); }
private:
    CRITICAL_SECTION *m_lock;
};

struct WRITE_PLUGIN
{
    HMODULE hDll;
    BOOL (WINAPI *pfnCreateCtrl)(DWORD*);
    BOOL (WINAPI *pfnDeleteCtrl)(DWORD);
    BOOL (WINAPI *pfnStartSave)(DWORD,LPCWSTR,BOOL,ULONGLONG);
    BOOL (WINAPI *pfnStopSave)(DWORD);
    BOOL (WINAPI *pfnGetSaveFilePath)(DWORD,WCHAR*,DWORD*);
    BOOL (WINAPI *pfnAddTSBuff)(DWORD,BYTE*,DWORD,DWORD*);
    bool Load(LPCWSTR path) {
        hDll = LoadLibraryW(path);
        if (hDll) {
            if ((pfnCreateCtrl = reinterpret_cast<BOOL (WINAPI*)(DWORD*)>(GetProcAddress(hDll, "CreateCtrl"))) != nullptr &&
                (pfnDeleteCtrl = reinterpret_cast<BOOL (WINAPI*)(DWORD)>(GetProcAddress(hDll, "DeleteCtrl"))) != nullptr &&
                (pfnStartSave = reinterpret_cast<BOOL (WINAPI*)(DWORD,LPCWSTR,BOOL,ULONGLONG)>(GetProcAddress(hDll, "StartSave"))) != nullptr &&
                (pfnStopSave = reinterpret_cast<BOOL (WINAPI*)(DWORD)>(GetProcAddress(hDll, "StopSave"))) != nullptr &&
                (pfnGetSaveFilePath = reinterpret_cast<BOOL (WINAPI*)(DWORD,WCHAR*,DWORD*)>(GetProcAddress(hDll, "GetSaveFilePath"))) != nullptr &&
                (pfnAddTSBuff = reinterpret_cast<BOOL (WINAPI*)(DWORD,BYTE*,DWORD,DWORD*)>(GetProcAddress(hDll, "AddTSBuff"))) != nullptr) {
                return true;
            }
            FreeLibrary(hDll);
            hDll = nullptr;
        }
        return false;
    }
};

struct WRITE_CONTEXT
{
    bool fUsed;
    bool fOverwrite;
    DWORD dataCount;
    DWORD listLen;
    bool fStartList[MAX_DUP];
    union {
        HANDLE fileList[MAX_DUP];
        DWORD ctrlList[MAX_DUP];
    };
    WCHAR savePath[MY_MAX_PATH];
};

namespace
{

HINSTANCE s_hinstDll;
CRITICAL_SECTION s_ctxLock;
DWORD s_dataUnit;
WRITE_PLUGIN s_writePlugin;
WRITE_CONTEXT s_ctxList[MAX_CONTEXT];

DWORD GetLongModuleFileName(HMODULE hModule, LPWSTR lpFileName, DWORD nSize)
{
    WCHAR longOrShortName[MY_MAX_PATH];
    DWORD nRet = GetModuleFileNameW(hModule, longOrShortName, MY_MAX_PATH);
    if (nRet && nRet < MY_MAX_PATH) {
        nRet = GetLongPathNameW(longOrShortName, lpFileName, nSize);
        if (nRet < nSize) {
            return nRet;
        }
    }
    return 0;
}

HANDLE CreateSaveFile(WCHAR *savePath, int savePathSize, bool fOverwrite)
{
    HANDLE file = CreateFileW(savePath, GENERIC_WRITE, FILE_SHARE_READ, nullptr,
                              fOverwrite ? CREATE_ALWAYS : CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE && lstrlenW(savePath) < MY_MAX_PATH) {
        int ext = lstrlenW(savePath) - 1;
        for (; ext >= 0 && savePath[ext] != L'\\' && savePath[ext] != L'/' && savePath[ext] != L'.'; --ext);
        ext = (ext >= 0 && savePath[ext] == L'.' ? ext : lstrlenW(savePath));
        WCHAR path[MY_MAX_PATH + 16];
        lstrcpyW(path, savePath);
        for (int i = 1; i < 1000; ++i) {
            wsprintfW(&path[ext], L"-(%d)%s", i, &savePath[ext]);
            if (lstrlenW(path) < savePathSize) {
                file = CreateFileW(path, GENERIC_WRITE, FILE_SHARE_READ, nullptr,
                                   fOverwrite ? CREATE_ALWAYS : CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
                if (file != INVALID_HANDLE_VALUE) {
                    lstrcpyW(savePath, path);
                    break;
                }
            }
        }
    }
    return file;
}

}

// PlugInの名前を取得する
DLL_EXPORT BOOL WINAPI GetPlugInName(WCHAR *name, DWORD *nameSize)
{
    static const WCHAR PLUGIN_NAME[] = L"Multiple Write PlugIn";
    static const DWORD size = sizeof(PLUGIN_NAME) / sizeof(PLUGIN_NAME[0]);
    if (nameSize) {
        if (!name) {
            *nameSize = size;
            return TRUE;
        }
        if (*nameSize >= size) {
            lstrcpyW(name, PLUGIN_NAME);
            return TRUE;
        }
        *nameSize = size;
    }
    return FALSE;
}

// PlugInで設定が必要な場合、設定用のダイアログなどを表示する
DLL_EXPORT void WINAPI Setting(HWND parentWnd)
{
    static_cast<void>(parentWnd);

    WCHAR path[MY_MAX_PATH];
    if (GetLongModuleFileName(s_hinstDll, path, MY_MAX_PATH - 4)) {
        lstrcatW(path, L".ini");
        WCHAR s[2];
        GetPrivateProfileStringW(L"SET", L"WritePlugin", L"*", s, 2, path);
        if (!lstrcmpiW(s, L"*")) {
            WritePrivateProfileStringW(L"SET", L"WritePlugin", L";Write_Default.dll", path);
            WritePrivateProfileStringW(L"SET", L"DataUnit", L"188", path);
        }
        ShellExecuteW(nullptr, L"edit", path, nullptr, nullptr, SW_SHOWNORMAL);
    }
}

// 複数保存対応のためインスタンスを新規に作成する
DLL_EXPORT BOOL WINAPI CreateCtrl(DWORD *id)
{
    CBlockLock lock(&s_ctxLock);
    for (DWORD i = 0; id && i < MAX_CONTEXT; ++i) {
        if (!s_ctxList[i].fUsed) {
            DWORD checkUsed = 0;
            for (; checkUsed < MAX_CONTEXT && !s_ctxList[checkUsed].fUsed; ++checkUsed);
            if (checkUsed == MAX_CONTEXT) {
                s_dataUnit = 1;
                s_writePlugin.hDll = nullptr;
                WCHAR path[MY_MAX_PATH];
                if (GetLongModuleFileName(s_hinstDll, path, MY_MAX_PATH - 4)) {
                    lstrcatW(path, L".ini");
                    s_dataUnit = GetPrivateProfileIntW(L"SET", L"DataUnit", 188, path);
                    s_dataUnit = max(s_dataUnit, 1);
                    WCHAR name[64];
                    GetPrivateProfileStringW(L"SET", L"WritePlugin", L"", name, 64, path);
                    if (name[0] && name[0] != L';') {
                        for (int j = lstrlenW(path) - 1; j >= 0 && path[j] != L'\\' && path[j] != L'/'; path[j--] = L'\0');
                        if (lstrlenW(path) + lstrlenW(name) < MY_MAX_PATH) {
                            lstrcatW(path, name);
                            s_writePlugin.Load(path);
                        }
                    }
                }
            }
            s_ctxList[i].fUsed = true;
            s_ctxList[i].listLen = 0;
            s_ctxList[i].savePath[0] = L'\0';
            *id = i + 1;
            return TRUE;
        }
    }
    return FALSE;
}

// ファイル保存を終了する
DLL_EXPORT BOOL WINAPI StopSave(DWORD id)
{
    CBlockLock lock(&s_ctxLock);
    if (--id < MAX_CONTEXT && s_ctxList[id].fUsed) {
        for (DWORD i = 0; i < s_ctxList[id].listLen; ++i) {
            if (s_writePlugin.hDll) {
                if (s_ctxList[id].ctrlList[i] != 0) {
                    s_writePlugin.pfnStopSave(s_ctxList[id].ctrlList[i]);
                    s_writePlugin.pfnDeleteCtrl(s_ctxList[id].ctrlList[i]);
                }
            }
            else {
                if (s_ctxList[id].fileList[i] != INVALID_HANDLE_VALUE) {
                    CloseHandle(s_ctxList[id].fileList[i]);
                }
            }
        }
        s_ctxList[id].listLen = 0;
        return TRUE;
    }
    return FALSE;
}

// ファイル保存を開始する
DLL_EXPORT BOOL WINAPI StartSave(DWORD id, LPCWSTR fileName, BOOL overWriteFlag, ULONGLONG createSize)
{
    CBlockLock lock(&s_ctxLock);
    if (StopSave(id--)) {
        WRITE_CONTEXT &ctx = s_ctxList[id];
        ctx.fOverwrite = !!overWriteFlag;
        ctx.dataCount = 0;
        ctx.fStartList[0] = false;
        if (s_writePlugin.hDll) {
            if (s_writePlugin.pfnCreateCtrl(&ctx.ctrlList[0])) {
                if (s_writePlugin.pfnStartSave(ctx.ctrlList[0], fileName, ctx.fOverwrite, createSize)) {
                    DWORD size = MY_MAX_PATH;
                    if (s_writePlugin.pfnGetSaveFilePath(ctx.ctrlList[0], ctx.savePath, &size)) {
                        ++ctx.listLen;
                        return TRUE;
                    }
                    s_writePlugin.pfnStopSave(ctx.ctrlList[0]);
                }
                s_writePlugin.pfnDeleteCtrl(ctx.ctrlList[0]);
            }
        }
        else {
            if (lstrlenW(fileName) < MY_MAX_PATH) {
                lstrcpyW(ctx.savePath, fileName);
                ctx.fileList[0] = CreateSaveFile(ctx.savePath, MY_MAX_PATH, ctx.fOverwrite);
                if (ctx.fileList[0] != INVALID_HANDLE_VALUE) {
                    ++ctx.listLen;
                    return TRUE;
                }
            }
        }
        ctx.savePath[0] = L'\0';
    }
    return FALSE;
}

// インスタンスを削除する
DLL_EXPORT BOOL WINAPI DeleteCtrl(DWORD id)
{
    CBlockLock lock(&s_ctxLock);
    if (StopSave(id--)) {
        s_ctxList[id].fUsed = false;
        DWORD checkUsed = 0;
        for (; checkUsed < MAX_CONTEXT && !s_ctxList[checkUsed].fUsed; ++checkUsed);
        if (checkUsed == MAX_CONTEXT) {
            if (s_writePlugin.hDll) {
                FreeLibrary(s_writePlugin.hDll);
            }
        }
        return TRUE;
    }
    return FALSE;
}

// 実際に保存しているファイルパスを取得する（再生やバッチ処理に利用される）
DLL_EXPORT BOOL WINAPI GetSaveFilePath(DWORD id, WCHAR *filePath, DWORD *filePathSize)
{
    CBlockLock lock(&s_ctxLock);
    if (--id < MAX_CONTEXT && s_ctxList[id].fUsed && filePathSize) {
        DWORD size = lstrlenW(s_ctxList[id].savePath) + 1;
        if (!filePath) {
            *filePathSize = size;
            return TRUE;
        }
        if (*filePathSize >= size) {
            lstrcpyW(filePath, s_ctxList[id].savePath);
            return TRUE;
        }
        *filePathSize = size;
    }
    return FALSE;
}

// 保存用データを送る
DLL_EXPORT BOOL WINAPI AddTSBuff(DWORD id, BYTE *data, DWORD size, DWORD *writeSize)
{
    CBlockLock lock(&s_ctxLock);
    if (--id < MAX_CONTEXT && s_ctxList[id].fUsed && data && size != 0 && writeSize) {
        WRITE_CONTEXT &ctx = s_ctxList[id];
        *writeSize = 0;
        bool fSuccessOneOrMore = false;
        for (DWORD i = 0; i < ctx.listLen; ++i) {
            if (s_writePlugin.hDll) {
                if (ctx.ctrlList[i] != 0) {
                    DWORD offset = (ctx.fStartList[i] || ctx.dataCount == 0 ? 0 : s_dataUnit - ctx.dataCount);
                    if (offset < size) {
                        ctx.fStartList[i] = true;
                        if (s_writePlugin.pfnAddTSBuff(ctx.ctrlList[i], data + offset, size - offset, writeSize)) {
                            fSuccessOneOrMore = true;
                        }
                        else {
                            s_writePlugin.pfnStopSave(ctx.ctrlList[i]);
                            s_writePlugin.pfnDeleteCtrl(ctx.ctrlList[i]);
                            ctx.ctrlList[i] = 0;
                        }
                    }
                }
            }
            else {
                if (ctx.fileList[i] != INVALID_HANDLE_VALUE) {
                    DWORD offset = (ctx.fStartList[i] || ctx.dataCount == 0 ? 0 : s_dataUnit - ctx.dataCount);
                    if (offset < size) {
                        ctx.fStartList[i] = true;
                        if (WriteFile(ctx.fileList[i], data + offset, size - offset, writeSize, nullptr)) {
                            fSuccessOneOrMore = true;
                        }
                        else {
                            CloseHandle(ctx.fileList[i]);
                            ctx.fileList[i] = INVALID_HANDLE_VALUE;
                        }
                    }
                }
            }
        }
        if (fSuccessOneOrMore) {
            ctx.dataCount = (ctx.dataCount + size) % s_dataUnit;
            *writeSize = size;
            return TRUE;
        }
    }
    return FALSE;
}

// ファイル保存を複製する
// 戻り値：
//  TRUE（保存開始/終了が実際に行われた）、FALSE（その他）
// 引数：
//  originalPath    [IN]GetSaveFilePath()で取得できるオリジナルの保存先パス
//  targetID        [IN(!targetPath)/OUT(targetPath)]保存ID(>0)。オリジナルの保存IDは必ず1
//  targetPath      [IN/OUT]保存先パス。成功すると実際の保存先パスと保存IDが格納される。!targetPathのときは保存終了
//  targetPathSize  [IN]targetPathのサイズ。最長6文字分("-(999)")の修飾を追加できる領域を確保すべき
//  overWriteFlag   [IN]負のときStartSave()のoverWriteFlagを継承、非負のときStartSave()と同義
//  createSize      [IN]StartSave()のcreateSizeと同義
DLL_EXPORT BOOL WINAPI DuplicateSave(LPCWSTR originalPath, DWORD *targetID, WCHAR *targetPath, DWORD targetPathSize, int overWriteFlag, ULONGLONG createSize)
{
    if (originalPath && targetID) {
        CBlockLock lock(&s_ctxLock);
        for (DWORD i = 0; i < MAX_CONTEXT; ++i) {
            WRITE_CONTEXT &ctx = s_ctxList[i];
            if (ctx.fUsed && !lstrcmpiW(ctx.savePath, originalPath)) {
                if (targetPath) {
                    // 保存開始
                    DWORD j = 0;
                    for (; j < ctx.listLen && (s_writePlugin.hDll && ctx.ctrlList[j] != 0 || !s_writePlugin.hDll && ctx.fileList[j] != INVALID_HANDLE_VALUE); ++j);
                    if (j >= ctx.listLen && j < MAX_DUP) {
                        ++ctx.listLen;
                    }
                    if (j < ctx.listLen) {
                        bool fOverwrite = overWriteFlag < 0 ? ctx.fOverwrite : !!overWriteFlag;
                        ctx.fStartList[j] = false;
                        if (s_writePlugin.hDll) {
                            if (s_writePlugin.pfnCreateCtrl(&ctx.ctrlList[j])) {
                                if (s_writePlugin.pfnStartSave(ctx.ctrlList[j], targetPath, fOverwrite, createSize)) {
                                    if (s_writePlugin.pfnGetSaveFilePath(ctx.ctrlList[j], targetPath, &targetPathSize)) {
                                        *targetID = j + 1;
                                        return TRUE;
                                    }
                                    s_writePlugin.pfnStopSave(ctx.ctrlList[j]);
                                }
                                s_writePlugin.pfnDeleteCtrl(ctx.ctrlList[j]);
                            }
                            ctx.ctrlList[j] = 0;
                        }
                        else {
                            ctx.fileList[j] = CreateSaveFile(targetPath, targetPathSize, fOverwrite);
                            if (ctx.fileList[j] != INVALID_HANDLE_VALUE) {
                                *targetID = j + 1;
                                return TRUE;
                            }
                        }
                    }
                }
                else {
                    // 保存終了
                    DWORD j = *targetID - 1;
                    if (j < ctx.listLen) {
                        if (s_writePlugin.hDll) {
                            if (ctx.ctrlList[j] != 0) {
                                s_writePlugin.pfnStopSave(ctx.ctrlList[j]);
                                s_writePlugin.pfnDeleteCtrl(ctx.ctrlList[j]);
                                ctx.ctrlList[j] = 0;
                                return TRUE;
                            }
                        }
                        else {
                            if (ctx.fileList[j] != INVALID_HANDLE_VALUE) {
                                CloseHandle(ctx.fileList[j]);
                                ctx.fileList[j] = INVALID_HANDLE_VALUE;
                                return TRUE;
                            }
                        }
                    }
                }
                break;
            }
        }
    }
    return FALSE;
}

BOOL WINAPI DllMain(HINSTANCE hinstDll, DWORD dwReason, LPVOID lpReserved)
{
    static_cast<void>(lpReserved);

    switch (dwReason) {
    case DLL_PROCESS_ATTACH:
        s_hinstDll = hinstDll;
        InitializeCriticalSection(&s_ctxLock);
        break;
    case DLL_PROCESS_DETACH:
        DeleteCriticalSection(&s_ctxLock);
        break;
    }
    return TRUE;
}
