#include "ToolbarSetup.h"

#include <aced.h>
#include <windows.h>
#include <oleauto.h>
#include <stdarg.h>

#include <fstream>
#include <string>

namespace {

const char* kToolbarName = "MG-Panel";
const char* kMenuGroupName = "MG-Project";

std::string Utf8ToAcp(const char* utf8) {
    if (utf8 == NULL || utf8[0] == '\0') return std::string();
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
    if (wlen <= 0) return std::string(utf8);
    std::wstring w;
    w.resize(wlen - 1);
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, &w[0], wlen);
    int alen = WideCharToMultiByte(CP_ACP, 0, w.c_str(), -1, NULL, 0, NULL, NULL);
    if (alen <= 0) return std::string(utf8);
    std::string a;
    a.resize(alen - 1);
    WideCharToMultiByte(CP_ACP, 0, w.c_str(), -1, &a[0], alen, NULL, NULL);
    return a;
}

HRESULT AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, LPOLESTR ptName, int cArgs = 0, ...) {
    if (!pDisp) return E_FAIL;

    DISPID dispID;
    HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_SYSTEM_DEFAULT, &dispID);
    if (FAILED(hr)) return hr;

    VARIANT* pArgs = new VARIANT[cArgs + 1];
    for (int i = 0; i < cArgs; ++i) VariantInit(&pArgs[i]);

    va_list marker;
    va_start(marker, cArgs);
    for (int i = 0; i < cArgs; i++) {
        pArgs[i] = va_arg(marker, VARIANT);
    }
    va_end(marker);

    DISPPARAMS dp = {NULL, NULL, 0, 0};
    dp.cArgs = cArgs;
    dp.rgvarg = pArgs;

    DISPID dispidNamed = DISPID_PROPERTYPUT;
    if (autoType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);

    delete[] pArgs;
    return hr;
}

std::wstring ToWide(const std::string& s) {
    if (s.empty()) {
        return std::wstring();
    }

    int n = MultiByteToWideChar(CP_ACP, 0, s.c_str(), -1, NULL, 0);
    if (n <= 0) {
        return std::wstring();
    }

    std::wstring ws;
    ws.resize(n - 1);
    MultiByteToWideChar(CP_ACP, 0, s.c_str(), -1, &ws[0], n);
    return ws;
}

bool EqualsNoCase(const std::string& a, const std::string& b) {
    if (a.size() != b.size()) {
        return false;
    }
    for (size_t i = 0; i < a.size(); ++i) {
        char ca = a[i];
        char cb = b[i];
        if (ca >= 'A' && ca <= 'Z') ca = (char)(ca - 'A' + 'a');
        if (cb >= 'A' && cb <= 'Z') cb = (char)(cb - 'A' + 'a');
        if (ca != cb) return false;
    }
    return true;
}

std::string ToAnsi(const BSTR s) {
    if (s == NULL) {
        return std::string();
    }

    int n = WideCharToMultiByte(CP_ACP, 0, s, -1, NULL, 0, NULL, NULL);
    if (n <= 0) {
        return std::string();
    }

    std::string out;
    out.resize(n - 1);
    WideCharToMultiByte(CP_ACP, 0, s, -1, &out[0], n, NULL, NULL);
    return out;
}

std::string ModuleDir() {
    HMODULE hm = NULL;
    MEMORY_BASIC_INFORMATION mbi;
    ZeroMemory(&mbi, sizeof(mbi));
    if (VirtualQuery((LPCVOID)&EnsureMgPanelToolbar, &mbi, sizeof(mbi)) != 0) {
        hm = (HMODULE)mbi.AllocationBase;
    }
    if (hm == NULL) {
        hm = GetModuleHandle(NULL);
    }

    char path[MAX_PATH];
    path[0] = '\0';
    GetModuleFileNameA(hm, path, MAX_PATH);

    std::string p(path);
    const std::string::size_type pos = p.find_last_of("\\/");
    if (pos == std::string::npos) {
        return std::string(".");
    }
    return p.substr(0, pos);
}

bool FileExists(const std::string& path) {
    DWORD attrs = GetFileAttributesA(path.c_str());
    return attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

IDispatch* GetToolbarByName(IDispatch* toolbars, const char* name) {
    // Fast path: direct access by name.
    VARIANT vName;
    VariantInit(&vName);
    vName.vt = VT_BSTR;
    std::wstring w = ToWide(name);
    vName.bstrVal = SysAllocString(w.c_str());

    VARIANT vToolbar;
    VariantInit(&vToolbar);
    HRESULT hr = AutoWrap(DISPATCH_METHOD, &vToolbar, toolbars, L"Item", 1, vName);
    SysFreeString(vName.bstrVal);
    if (SUCCEEDED(hr) && vToolbar.vt == VT_DISPATCH && vToolbar.pdispVal != NULL) {
        return vToolbar.pdispVal;
    }
    VariantClear(&vToolbar);

    // Fallback: enumerate and compare Name case-insensitively.
    VARIANT vCount;
    VariantInit(&vCount);
    if (FAILED(AutoWrap(DISPATCH_PROPERTYGET, &vCount, toolbars, L"Count", 0))) {
        return NULL;
    }

    long count = 0;
    if (vCount.vt == VT_I4) {
        count = vCount.lVal;
    } else if (vCount.vt == VT_I2) {
        count = vCount.iVal;
    }
    VariantClear(&vCount);

    for (long i = 0; i < count; ++i) {
        VARIANT vIndex;
        VariantInit(&vIndex);
        vIndex.vt = VT_I4;
        vIndex.lVal = i;

        VARIANT vTb;
        VariantInit(&vTb);
        if (FAILED(AutoWrap(DISPATCH_METHOD, &vTb, toolbars, L"Item", 1, vIndex)) ||
            vTb.vt != VT_DISPATCH || vTb.pdispVal == NULL) {
            VariantClear(&vTb);
            continue;
        }

        IDispatch* tb = vTb.pdispVal;
        VARIANT vTbName;
        VariantInit(&vTbName);
        if (SUCCEEDED(AutoWrap(DISPATCH_PROPERTYGET, &vTbName, tb, L"Name", 0)) && vTbName.vt == VT_BSTR) {
            const std::string tbName = ToAnsi(vTbName.bstrVal);
            if (EqualsNoCase(tbName, name)) {
                VariantClear(&vTbName);
                return tb;
            }
        }
        VariantClear(&vTbName);
        tb->Release();
    }

    return NULL;
}

IDispatch* GetMenuGroupByName(IDispatch* menuGroups, const char* name) {
    VARIANT vCount;
    VariantInit(&vCount);
    if (FAILED(AutoWrap(DISPATCH_PROPERTYGET, &vCount, menuGroups, L"Count", 0))) {
        return NULL;
    }

    long count = 0;
    if (vCount.vt == VT_I4) count = vCount.lVal;
    else if (vCount.vt == VT_I2) count = vCount.iVal;
    VariantClear(&vCount);

    for (long i = 0; i < count; ++i) {
        VARIANT vIndex;
        VariantInit(&vIndex);
        vIndex.vt = VT_I4;
        vIndex.lVal = i;

        VARIANT vMenuGroup;
        VariantInit(&vMenuGroup);
        if (FAILED(AutoWrap(DISPATCH_METHOD, &vMenuGroup, menuGroups, L"Item", 1, vIndex)) ||
            vMenuGroup.vt != VT_DISPATCH || vMenuGroup.pdispVal == NULL) {
            VariantClear(&vMenuGroup);
            continue;
        }

        IDispatch* menuGroup = vMenuGroup.pdispVal;
        VARIANT vName;
        VariantInit(&vName);
        if (SUCCEEDED(AutoWrap(DISPATCH_PROPERTYGET, &vName, menuGroup, L"Name", 0)) && vName.vt == VT_BSTR) {
            const std::string groupName = ToAnsi(vName.bstrVal);
            if (EqualsNoCase(groupName, name)) {
                VariantClear(&vName);
                return menuGroup;
            }
        }
        VariantClear(&vName);
        menuGroup->Release();
    }

    return NULL;
}

IDispatch* LoadMenuGroup(IDispatch* menuGroups, const std::string& menuFilePath) {
    VARIANT vPath;
    VariantInit(&vPath);
    vPath.vt = VT_BSTR;
    vPath.bstrVal = SysAllocString(ToWide(menuFilePath).c_str());

    VARIANT vBaseMenu;
    VariantInit(&vBaseMenu);
    vBaseMenu.vt = VT_BOOL;
    vBaseMenu.boolVal = VARIANT_FALSE;

    VARIANT vLoaded;
    VariantInit(&vLoaded);
    const HRESULT hr = AutoWrap(DISPATCH_METHOD, &vLoaded, menuGroups, L"Load", 2, vBaseMenu, vPath);
    SysFreeString(vPath.bstrVal);

    if (FAILED(hr) || vLoaded.vt != VT_DISPATCH || vLoaded.pdispVal == NULL) {
        VariantClear(&vLoaded);
        return NULL;
    }

    return vLoaded.pdispVal;
}

IDispatch* EnsureMgProjectMenuGroup(IDispatch* menuGroups) {
    const std::string dir = ModuleDir();
    const std::string cuixPath = dir + "\\MG-Project.cuix";
    const std::string cuiPath = dir + "\\MG-Project.cui";
    const std::string mnsPath = dir + "\\MG-Project.mns";

    // First-run bootstrap: create a basic MNS menu definition if no menu file exists yet.
    if (!FileExists(cuixPath) && !FileExists(cuiPath) && !FileExists(mnsPath)) {
        std::ofstream mns(mnsPath.c_str(), std::ios::out | std::ios::trunc);
        if (mns) {
            mns << "***MENUGROUP=" << kMenuGroupName << "\n";
            mns << "***TOOLBARS\n";
            mns.close();
        }
    }

    IDispatch* mgGroup = GetMenuGroupByName(menuGroups, kMenuGroupName);
    if (mgGroup != NULL) {
        return mgGroup;
    }

    const std::string candidates[] = {
        dir + "\\MG-Project.cuix",
        dir + "\\MG-Project.cui",
        dir + "\\MG-Project.mns"
    };

    for (size_t i = 0; i < sizeof(candidates) / sizeof(candidates[0]); ++i) {
        if (!FileExists(candidates[i])) {
            continue;
        }

        IDispatch* loaded = LoadMenuGroup(menuGroups, candidates[i]);
        if (loaded != NULL) {
            return loaded;
        }
    }

    return NULL;
}

IDispatch* AddToolbar(IDispatch* toolbars, const char* name) {
    VARIANT vName;
    VariantInit(&vName);
    vName.vt = VT_BSTR;
    std::wstring w = ToWide(name);
    vName.bstrVal = SysAllocString(w.c_str());

    VARIANT vToolbar;
    VariantInit(&vToolbar);
    HRESULT hr = AutoWrap(DISPATCH_METHOD, &vToolbar, toolbars, L"Add", 1, vName);
    SysFreeString(vName.bstrVal);
    if (FAILED(hr) || vToolbar.vt != VT_DISPATCH || vToolbar.pdispVal == NULL) {
        VariantClear(&vToolbar);
        return NULL;
    }
    return vToolbar.pdispVal;
}

void AddButton(IDispatch* toolbar,
               int index,
               const char* name,
               const char* help,
               const char* macro,
               const char* smallBmpPath,
               const char* largeBmpPath) {
    VARIANT vIndex;
    VariantInit(&vIndex);
    vIndex.vt = VT_I4;
    vIndex.lVal = index;

    VARIANT vName;
    VariantInit(&vName);
    vName.vt = VT_BSTR;
    vName.bstrVal = SysAllocString(ToWide(name).c_str());

    VARIANT vHelp;
    VariantInit(&vHelp);
    vHelp.vt = VT_BSTR;
    vHelp.bstrVal = SysAllocString(ToWide(help).c_str());

    VARIANT vMacro;
    VariantInit(&vMacro);
    vMacro.vt = VT_BSTR;
    vMacro.bstrVal = SysAllocString(ToWide(macro).c_str());

    VARIANT vBtn;
    VariantInit(&vBtn);
    if (SUCCEEDED(AutoWrap(DISPATCH_METHOD, &vBtn, toolbar, L"AddToolbarButton", 4, vMacro, vHelp, vName, vIndex)) &&
        vBtn.vt == VT_DISPATCH && vBtn.pdispVal != NULL) {
        IDispatch* btn = vBtn.pdispVal;

        if (smallBmpPath != NULL && largeBmpPath != NULL &&
            FileExists(smallBmpPath) && FileExists(largeBmpPath)) {
            VARIANT vSmall;
            VariantInit(&vSmall);
            vSmall.vt = VT_BSTR;
            vSmall.bstrVal = SysAllocString(ToWide(smallBmpPath).c_str());

            VARIANT vLarge;
            VariantInit(&vLarge);
            vLarge.vt = VT_BSTR;
            vLarge.bstrVal = SysAllocString(ToWide(largeBmpPath).c_str());

            AutoWrap(DISPATCH_METHOD, NULL, btn, L"SetBitmaps", 2, vLarge, vSmall);

            SysFreeString(vSmall.bstrVal);
            SysFreeString(vLarge.bstrVal);
        }

        btn->Release();
    }

    SysFreeString(vName.bstrVal);
    SysFreeString(vHelp.bstrVal);
    SysFreeString(vMacro.bstrVal);
}

void SetButtonBitmaps(IDispatch* toolbar, int index, const char* smallBmpPath, const char* largeBmpPath) {
    if (toolbar == NULL || smallBmpPath == NULL || largeBmpPath == NULL) {
        return;
    }
    if (!FileExists(smallBmpPath) || !FileExists(largeBmpPath)) {
        return;
    }

    VARIANT vIndex;
    VariantInit(&vIndex);
    vIndex.vt = VT_I4;
    vIndex.lVal = index;

    VARIANT vBtn;
    VariantInit(&vBtn);
    if (FAILED(AutoWrap(DISPATCH_METHOD, &vBtn, toolbar, L"Item", 1, vIndex)) ||
        vBtn.vt != VT_DISPATCH || vBtn.pdispVal == NULL) {
        VariantClear(&vBtn);
        return;
    }

    IDispatch* btn = vBtn.pdispVal;

    VARIANT vSmall;
    VariantInit(&vSmall);
    vSmall.vt = VT_BSTR;
    vSmall.bstrVal = SysAllocString(ToWide(smallBmpPath).c_str());

    VARIANT vLarge;
    VariantInit(&vLarge);
    vLarge.vt = VT_BSTR;
    vLarge.bstrVal = SysAllocString(ToWide(largeBmpPath).c_str());

    AutoWrap(DISPATCH_METHOD, NULL, btn, L"SetBitmaps", 2, vLarge, vSmall);

    SysFreeString(vSmall.bstrVal);
    SysFreeString(vLarge.bstrVal);
    btn->Release();
}

}  // namespace

void EnsureMgPanelToolbar() {
    IDispatch* acad = acedGetIDispatch(TRUE);
    if (acad == NULL) {
        return;
    }

    VARIANT vMenuGroups;
    VariantInit(&vMenuGroups);
    if (FAILED(AutoWrap(DISPATCH_PROPERTYGET, &vMenuGroups, acad, L"MenuGroups", 0)) ||
        vMenuGroups.vt != VT_DISPATCH || vMenuGroups.pdispVal == NULL) {
        acad->Release();
        return;
    }

    IDispatch* menuGroups = vMenuGroups.pdispVal;
    IDispatch* menuGroup = EnsureMgProjectMenuGroup(menuGroups);
    if (menuGroup == NULL) {
        acutPrintf("\n%s", Utf8ToAcp("MG-Panel: группа меню не загружена и файл меню не найден.").c_str());
        menuGroups->Release();
        acad->Release();
        return;
    }

    VARIANT vToolbars;
    VariantInit(&vToolbars);
    if (FAILED(AutoWrap(DISPATCH_PROPERTYGET, &vToolbars, menuGroup, L"Toolbars", 0)) ||
        vToolbars.vt != VT_DISPATCH || vToolbars.pdispVal == NULL) {
        menuGroup->Release();
        menuGroups->Release();
        acad->Release();
        return;
    }

    IDispatch* toolbars = vToolbars.pdispVal;
    IDispatch* toolbar = GetToolbarByName(toolbars, kToolbarName);
    bool createdToolbar = false;
    const std::string dir = ModuleDir();
    const std::string show16 = dir + "\\Loadpanel16.bmp";
    const std::string show32 = dir + "\\Loadpanel32.bmp";
    const std::string hide16 = dir + "\\UnLoadpanel16.bmp";
    const std::string hide32 = dir + "\\UnLoadpanel32.bmp";
    if (toolbar == NULL) {
        toolbar = AddToolbar(toolbars, kToolbarName);
        if (toolbar != NULL) {
            createdToolbar = true;
            const std::string showName = Utf8ToAcp("ПОКАЗ");
            const std::string showHelp = Utf8ToAcp("Показать панель свойств DWG");
            const std::string hideName = Utf8ToAcp("СКРЫТЬ");
            const std::string hideHelp = Utf8ToAcp("Скрыть панель свойств DWG");
            AddButton(toolbar, 0, showName.c_str(), showHelp.c_str(), "SHOWDWGPROPS ", show16.c_str(), show32.c_str());
            AddButton(toolbar, 1, hideName.c_str(), hideHelp.c_str(), "HIDEDWGPROPS ", hide16.c_str(), hide32.c_str());
        }
    }
    if (toolbar != NULL) {
        SetButtonBitmaps(toolbar, 0, show16.c_str(), show32.c_str());
        SetButtonBitmaps(toolbar, 1, hide16.c_str(), hide32.c_str());
    }

    if (toolbar != NULL) {
        VARIANT vVisible;
        VariantInit(&vVisible);
        vVisible.vt = VT_BOOL;
        vVisible.boolVal = VARIANT_TRUE;
        AutoWrap(DISPATCH_PROPERTYPUT, NULL, toolbar, L"Visible", 1, vVisible);
        toolbar->Release();
    }

    // Persist toolbar placement/settings in the MG-Project menu group so it is
    // not recreated with default position on each AutoCAD start.
    if (createdToolbar) {
        AutoWrap(DISPATCH_METHOD, NULL, menuGroup, L"Save", 0);
    }

    toolbars->Release();
    menuGroup->Release();
    menuGroups->Release();
    acad->Release();
}
