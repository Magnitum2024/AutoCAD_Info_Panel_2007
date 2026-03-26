#include "DwgPropsPanel.h"
#include "Xlsx2DwgProp.h"

#include <aced.h>
#include <adslib.h>
#include <dbapserv.h>
#include <summinfo.h>
#include <string.h>
#include <stdarg.h>

namespace {

std::string ToLocal8(const ACHAR* text) {
    if (text == NULL) {
        return "";
    }

#ifdef AD_UNICODE
    const int sizeNeeded = WideCharToMultiByte(CP_ACP, 0, text, -1, NULL, 0, NULL, NULL);
    if (sizeNeeded <= 0) {
        return "";
    }

    std::string out;
    out.resize(sizeNeeded - 1);
    WideCharToMultiByte(CP_ACP, 0, text, -1, &out[0], sizeNeeded, NULL, NULL);
    return out;
#else
    return std::string(text);
#endif
}

std::basic_string<ACHAR> AcpToAChar(const char* text) {
    if (text == NULL || text[0] == '\0') {
        return std::basic_string<ACHAR>();
    }
#ifdef AD_UNICODE
    int wlen = MultiByteToWideChar(CP_ACP, 0, text, -1, NULL, 0);
    if (wlen <= 0) {
        return std::basic_string<ACHAR>();
    }
    std::basic_string<ACHAR> w;
    w.resize(wlen - 1);
    MultiByteToWideChar(CP_ACP, 0, text, -1, &w[0], wlen);
    return w;
#else
    return std::basic_string<ACHAR>(text);
#endif
}

std::string Utf8ToAcp(const char* utf8) {
    if (utf8 == NULL || utf8[0] == '\0') {
        return std::string();
    }

    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
    if (wlen <= 0) {
        return std::string(utf8);
    }

    std::wstring w;
    w.resize(wlen - 1);
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, &w[0], wlen);

    int alen = WideCharToMultiByte(CP_ACP, 0, w.c_str(), -1, NULL, 0, NULL, NULL);
    if (alen <= 0) {
        return std::string(utf8);
    }

    std::string a;
    a.resize(alen - 1);
    WideCharToMultiByte(CP_ACP, 0, w.c_str(), -1, &a[0], alen, NULL, NULL);
    return a;
}

void PrintUtf8(const char* fmtUtf8, ...) {
    std::string fmt = Utf8ToAcp(fmtUtf8);

    char out[1024];
    out[0] = '\0';

    va_list args;
    va_start(args, fmtUtf8);
    _vsnprintf(out, sizeof(out) - 1, fmt.c_str(), args);
    va_end(args);

    out[sizeof(out) - 1] = '\0';
    const std::basic_string<ACHAR> msg = AcpToAChar(out);
    acutPrintf(ACRX_T("%s"), msg.c_str());
}

std::string UserConfigDir() {
    char buf[MAX_PATH];
    buf[0] = '\0';

    DWORD n = GetEnvironmentVariableA("APPDATA", buf, MAX_PATH);
    if (n == 0 || n >= MAX_PATH) {
        n = GetEnvironmentVariableA("USERPROFILE", buf, MAX_PATH);
        if (n == 0 || n >= MAX_PATH) {
            GetTempPathA(MAX_PATH, buf);
        }
    }

    std::string dir(buf);
    dir += "\\AutoCAD_Info_Panel";
    CreateDirectoryA(dir.c_str(), NULL);
    return dir;
}

std::string BoundsIniPath() {
    const std::string ini = UserConfigDir() + "\\DwgPropsPanel.ini";
    if (GetFileAttributesA(ini.c_str()) == INVALID_FILE_ATTRIBUTES) {
        WritePrivateProfileStringA("Panel", "X", "150", ini.c_str());
        WritePrivateProfileStringA("Panel", "Y", "150", ini.c_str());
        WritePrivateProfileStringA("Panel", "W", "420", ini.c_str());
        WritePrivateProfileStringA("Panel", "H", "560", ini.c_str());
    }
    return ini;
}

}  // namespace

CDwgPropsPanel::CDwgPropsPanel() : m_hWnd(NULL), m_hList(NULL), m_hBtnLoadXlsx(NULL), m_hBtnInsertField(NULL), m_hLblTrackedXlsx(NULL), m_hEditSearch(NULL), m_hComboGroup(NULL), m_fontNameBold(NULL), m_fontDesc(NULL), m_fontValue(NULL), m_isTopMost(false), m_trackedXlsxHashMismatch(false), m_lastTrackedXlsxCheckTick(0), m_lastTextInputHwnd(NULL) {}

CDwgPropsPanel::~CDwgPropsPanel() {
    Destroy();
}

bool CDwgPropsPanel::Create() {
    if (m_hWnd != NULL) {
        return true;
    }

    WNDCLASSA wc;
    ZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc = CDwgPropsPanel::WndProc;
    wc.hInstance = GetModuleHandle(NULL);
    wc.lpszClassName = "DWGPropsPanelWnd";
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);

    RegisterClassA(&wc);

    HWND owner = GetForegroundWindow();

    int x = 150;
    int y = 150;
    int w = 420;
    int h = 560;
    LoadSavedBounds(x, y, w, h);

    const std::string wndTitle = Utf8ToAcp("Свойства DWG");
    m_hWnd = CreateWindowExA(
        WS_EX_TOOLWINDOW,
        wc.lpszClassName,
        wndTitle.c_str(),
        WS_OVERLAPPEDWINDOW,
        x,
        y,
        w,
        h,
        owner,
        NULL,
        wc.hInstance,
        this);

    if (m_hWnd == NULL) {
        return false;
    }

    Show();
    ReloadProperties();
    return true;
}

void CDwgPropsPanel::Destroy() {
    if (m_hWnd != NULL) {
        SaveCurrentBounds();
        DestroyWindow(m_hWnd);
        m_hWnd = NULL;
        m_hList = NULL;
        m_hBtnLoadXlsx = NULL;
        m_hBtnInsertField = NULL;
        m_hLblTrackedXlsx = NULL;
        m_hEditSearch = NULL;
        m_hComboGroup = NULL;
        m_lastTextInputHwnd = NULL;
        if (m_fontNameBold != NULL) { DeleteObject(m_fontNameBold); m_fontNameBold = NULL; }
        if (m_fontDesc != NULL) { DeleteObject(m_fontDesc); m_fontDesc = NULL; }
        if (m_fontValue != NULL) { DeleteObject(m_fontValue); m_fontValue = NULL; }
    }
}

bool CDwgPropsPanel::LoadSavedBounds(int& x, int& y, int& w, int& h) const {
    const std::string ini = BoundsIniPath();
    const int savedX = GetPrivateProfileIntA("Panel", "X", x, ini.c_str());
    const int savedY = GetPrivateProfileIntA("Panel", "Y", y, ini.c_str());
    const int savedW = GetPrivateProfileIntA("Panel", "W", w, ini.c_str());
    const int savedH = GetPrivateProfileIntA("Panel", "H", h, ini.c_str());

    if (savedW < 260 || savedH < 220) {
        return false;
    }

    x = savedX;
    y = savedY;
    w = savedW;
    h = savedH;
    return true;
}

void CDwgPropsPanel::SaveCurrentBounds() const {
    if (m_hWnd == NULL) {
        return;
    }

    RECT rc;
    if (!GetWindowRect(m_hWnd, &rc)) {
        return;
    }

    char buf[64];
    const std::string ini = BoundsIniPath();

    _snprintf(buf, sizeof(buf) - 1, "%d", rc.left);
    buf[sizeof(buf) - 1] = '\0';
    WritePrivateProfileStringA("Panel", "X", buf, ini.c_str());

    _snprintf(buf, sizeof(buf) - 1, "%d", rc.top);
    buf[sizeof(buf) - 1] = '\0';
    WritePrivateProfileStringA("Panel", "Y", buf, ini.c_str());

    _snprintf(buf, sizeof(buf) - 1, "%d", rc.right - rc.left);
    buf[sizeof(buf) - 1] = '\0';
    WritePrivateProfileStringA("Panel", "W", buf, ini.c_str());

    _snprintf(buf, sizeof(buf) - 1, "%d", rc.bottom - rc.top);
    buf[sizeof(buf) - 1] = '\0';
    WritePrivateProfileStringA("Panel", "H", buf, ini.c_str());
}

void CDwgPropsPanel::Show() {
    if (m_hWnd != NULL) {
        RefreshTrackedXlsxState(true);
        ShowWindow(m_hWnd, SW_SHOW);
        SetForegroundWindow(m_hWnd);
        UpdateTopMostState();
    }
}

void CDwgPropsPanel::RefreshTrackedXlsxState(bool force) {
    if (m_hWnd == NULL || m_hLblTrackedXlsx == NULL) {
        return;
    }

    const int everyMinutes = Xlsx2DwgProp_GetHashCheckMinutes();
    const DWORD everyMs = (DWORD)everyMinutes * 60UL * 1000UL;
    const DWORD now = GetTickCount();
    if (!force && (now - m_lastTrackedXlsxCheckTick) < everyMs) {
        return;
    }
    m_lastTrackedXlsxCheckTick = now;

    std::string fileName;
    std::string fullPath;
    bool mismatch = false;
    if (!Xlsx2DwgProp_GetTrackedFileStatus(fileName, fullPath, mismatch)) {
        SetWindowTextA(m_hLblTrackedXlsx, "");
        m_trackedXlsxHashMismatch = false;
        InvalidateRect(m_hLblTrackedXlsx, NULL, TRUE);
        return;
    }

    std::string label = Utf8ToAcp("XLSX-файл: ");
    label += fileName;
    SetWindowTextA(m_hLblTrackedXlsx, label.c_str());
    m_trackedXlsxHashMismatch = mismatch;
    InvalidateRect(m_hLblTrackedXlsx, NULL, TRUE);
}

void CDwgPropsPanel::Hide() {
    if (m_hWnd != NULL) {
        ShowWindow(m_hWnd, SW_HIDE);
    }
}

void CDwgPropsPanel::AddProperty(const char* name, const std::string& value, bool isCustom, const std::string& description, const std::string& group) {
    ListItem item;
    item.key = name;
    item.value = value;
    item.isCustom = isCustom;
    item.description = description;
    item.group = group;
    m_allItems.push_back(item);
}

void CDwgPropsPanel::RebuildGroupFilter() {
    if (m_hComboGroup == NULL) return;

    const std::string allGroupsText = Utf8ToAcp("Все группы");
    SendMessageA(m_hComboGroup, CB_RESETCONTENT, 0, 0);
    SendMessageA(m_hComboGroup, CB_ADDSTRING, 0, (LPARAM)allGroupsText.c_str());

    std::vector<std::string> groups;
    Xlsx2DwgProp_GetGroups(groups);
    for (size_t i = 0; i < groups.size(); ++i) {
        if (!groups[i].empty()) {
            SendMessageA(m_hComboGroup, CB_ADDSTRING, 0, (LPARAM)groups[i].c_str());
        }
    }

    SendMessageA(m_hComboGroup, CB_SETCURSEL, 0, 0);
}

void CDwgPropsPanel::RebuildVisibleList() {
    if (m_hList == NULL) {
        return;
    }

    char filterBuf[256];
    filterBuf[0] = '\0';
    if (m_hEditSearch != NULL) {
        GetWindowTextA(m_hEditSearch, filterBuf, sizeof(filterBuf) - 1);
    }
    std::string filter = filterBuf;
    for (size_t i = 0; i < filter.size(); ++i) {
        if (filter[i] >= 'A' && filter[i] <= 'Z') {
            filter[i] = (char)(filter[i] - 'A' + 'a');
        }
    }

    SendMessageA(m_hList, LB_RESETCONTENT, 0, 0);
    m_items.clear();

    char groupBuf[256];
    groupBuf[0] = '\0';
    if (m_hComboGroup != NULL) {
        const int sel = (int)SendMessageA(m_hComboGroup, CB_GETCURSEL, 0, 0);
        if (sel >= 0) {
            SendMessageA(m_hComboGroup, CB_GETLBTEXT, sel, (LPARAM)groupBuf);
        }
    }
    const std::string selectedGroup = groupBuf;
    const std::string allGroupsText = Utf8ToAcp("Все группы");
    const bool useGroupFilter = !selectedGroup.empty() && selectedGroup != allGroupsText;

    for (size_t i = 0; i < m_allItems.size(); ++i) {
        const ListItem& src = m_allItems[i];
        if (useGroupFilter) {
            if (src.group.empty() || src.group != selectedGroup) {
                continue;
            }
        }
        std::string hay = src.key + " " + src.description + " " + src.value;
        for (size_t j = 0; j < hay.size(); ++j) {
            if (hay[j] >= 'A' && hay[j] <= 'Z') {
                hay[j] = (char)(hay[j] - 'A' + 'a');
            }
        }

        if (!filter.empty() && hay.find(filter) == std::string::npos) {
            continue;
        }

        m_items.push_back(src);
        std::string line(src.key);
        line += "	";
        line += src.value;
        SendMessageA(m_hList, LB_ADDSTRING, 0, (LPARAM)line.c_str());
    }
}

std::string CDwgPropsPanel::ReadSysVarAsString(const char* varName) const {
    resbuf rb;
    const std::basic_string<ACHAR> aVarName = AcpToAChar(varName);
    if (acedGetVar(aVarName.c_str(), &rb) != RTNORM) {
        return Utf8ToAcp("<н/д>");
    }

    char buf[128];
    buf[0] = '\0';

    switch (rb.restype) {
        case RTSHORT:
            _snprintf(buf, sizeof(buf) - 1, "%d", rb.resval.rint);
            buf[sizeof(buf) - 1] = '\0';
            break;
        case RTLONG:
            _snprintf(buf, sizeof(buf) - 1, "%ld", rb.resval.rlong);
            buf[sizeof(buf) - 1] = '\0';
            break;
        case RTREAL:
            _snprintf(buf, sizeof(buf) - 1, "%.3f", rb.resval.rreal);
            buf[sizeof(buf) - 1] = '\0';
            break;
        case RTSTR: {
            std::string s = ToLocal8(rb.resval.rstring);
            if (rb.resval.rstring != NULL) {
                acdbFree(rb.resval.rstring);
            }
            return s;
        }
        default:
            return Utf8ToAcp("<не поддерживается>");
    }

    return std::string(buf);
}

void CDwgPropsPanel::ReloadProperties() {
    if (m_hList == NULL) {
        return;
    }

    SendMessageA(m_hList, LB_RESETCONTENT, 0, 0);
    m_allItems.clear();
    m_items.clear();

    AddProperty("DWGNAME", ReadSysVarAsString("DWGNAME"), false);
    AddProperty("DWGPREFIX", ReadSysVarAsString("DWGPREFIX"), false);
    AddProperty("CTAB", ReadSysVarAsString("CTAB"), false);
    AddProperty("CDATE", ReadSysVarAsString("CDATE"), false);

    AcDbDatabase* pDb = acdbHostApplicationServices()->workingDatabase();
    if (pDb == NULL) {
        return;
    }

    AcDbDatabaseSummaryInfo* pInfo = NULL;
    if (acdbGetSummaryInfo(pDb, pInfo) == Acad::eOk && pInfo != NULL) {
        const int customCount = pInfo->numCustomInfo();
        if (customCount > 0) {
            AddProperty("---------------- ПОЛЬЗОВАТЕЛЬСКИЕ СВОЙСТВА DWG ----------------", "", false);
        }

        for (int i = 0; i < customCount; ++i) {
            ACHAR* key = NULL;
            ACHAR* val = NULL;
            if (pInfo->getCustomSummaryInfo(i, key, val) == Acad::eOk) {
                std::string desc;
                const std::string keyLocal = ToLocal8(key);
                Xlsx2DwgProp_GetDescriptionForKey(keyLocal, desc);
                std::string group;
                Xlsx2DwgProp_GetGroupForKey(keyLocal, group);
                AddProperty(keyLocal.c_str(), ToLocal8(val), true, desc, group);
            }
            if (key != NULL) {
                acdbFree(key);
            }
            if (val != NULL) {
                acdbFree(val);
            }
        }

        delete pInfo;
    }

    RebuildGroupFilter();
    RebuildVisibleList();
}


bool CDwgPropsPanel::IsHostAppForeground() const {
    HWND fg = GetForegroundWindow();
    if (fg == NULL) {
        return false;
    }

    DWORD pid = 0;
    GetWindowThreadProcessId(fg, &pid);
    return pid == GetCurrentProcessId();
}

void CDwgPropsPanel::UpdateTopMostState() {
    if (m_hWnd == NULL) {
        return;
    }

    const bool shouldBeTopMost = IsHostAppForeground();
    if (shouldBeTopMost == m_isTopMost) {
        return;
    }

    SetWindowPos(
        m_hWnd,
        shouldBeTopMost ? HWND_TOPMOST : HWND_NOTOPMOST,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

    m_isTopMost = shouldBeTopMost;
}


void CDwgPropsPanel::DrawListItem(const DRAWITEMSTRUCT* dis) {
    if (dis == NULL || dis->itemID == (UINT)-1) {
        return;
    }

    if (dis->itemID >= m_items.size()) {
        return;
    }

    const ListItem& li = m_items[dis->itemID];
    const std::string& key = li.key;
    const std::string& val = li.value;
    const std::string& desc = li.description;

    HDC hdc = dis->hDC;
    RECT rc = dis->rcItem;

    const bool selected = (dis->itemState & ODS_SELECTED) != 0;
    COLORREF bk = selected ? GetSysColor(COLOR_HIGHLIGHT) : GetSysColor(COLOR_WINDOW);
    COLORREF tx = selected ? GetSysColor(COLOR_HIGHLIGHTTEXT) : RGB(0, 0, 0);

    HBRUSH b = CreateSolidBrush(bk);
    FillRect(hdc, &rc, b);
    DeleteObject(b);

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, tx);

    RECT rcKey = rc;
    rcKey.left += 6;
    rcKey.top += 2;
    rcKey.bottom = rcKey.top + 18;
    rcKey.right = rc.right - 6;

    const int fullW = rcKey.right - rcKey.left;
    int keyColW = (fullW * 45) / 100;
    if (keyColW < 110) keyColW = 110;
    if (keyColW > 260) keyColW = 260;
    if (keyColW > fullW) keyColW = fullW;
    rcKey.right = rcKey.left + keyColW;

    RECT rcDesc = rcKey;
    rcDesc.left = rcKey.right + 8;
    rcDesc.right = rc.right - 6;

    RECT rcVal = rc;
    rcVal.left += 18;
    rcVal.top = rcKey.bottom;
    rcVal.bottom -= 2;

    HFONT old = (HFONT)SelectObject(hdc, m_fontNameBold != NULL ? m_fontNameBold : GetStockObject(DEFAULT_GUI_FONT));
    DrawTextA(hdc, key.c_str(), -1, &rcKey, DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_END_ELLIPSIS);

    if (!desc.empty() && rcDesc.left < rcDesc.right) {
        SelectObject(hdc, m_fontDesc != NULL ? m_fontDesc : GetStockObject(DEFAULT_GUI_FONT));
        DrawTextA(hdc, desc.c_str(), -1, &rcDesc, DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_END_ELLIPSIS);
    }

    SelectObject(hdc, m_fontValue != NULL ? m_fontValue : GetStockObject(DEFAULT_GUI_FONT));
    DrawTextA(hdc, val.c_str(), -1, &rcVal, DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_END_ELLIPSIS);

    SelectObject(hdc, old);

    if ((dis->itemState & ODS_FOCUS) != 0) {
        DrawFocusRect(hdc, &dis->rcItem);
    }
}

bool CDwgPropsPanel::IsTextInputActive() const {
    resbuf rb;
    if (acedGetVar(ACRX_T("CMDNAMES"), &rb) != RTNORM || rb.restype != RTSTR || rb.resval.rstring == NULL) {
        return false;
    }

    std::string cmd = ToLocal8(rb.resval.rstring);
    acdbFree(rb.resval.rstring);

    for (size_t i = 0; i < cmd.size(); ++i) {
        if (cmd[i] >= 'a' && cmd[i] <= 'z') cmd[i] = (char)(cmd[i] - 32);
    }

    return cmd.find("MTEXT") != std::string::npos || cmd.find("TEXT") != std::string::npos || cmd.find("MTEDIT") != std::string::npos || cmd.find("DDEDIT") != std::string::npos || cmd.find("TEXTEDIT") != std::string::npos;
}

bool CDwgPropsPanel::CopyTextToClipboard(const std::string& text) {
    if (text.empty()) return false;
    if (!OpenClipboard(m_hWnd)) return false;

    EmptyClipboard();
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, text.size() + 1);
    if (hMem == NULL) {
        CloseClipboard();
        return false;
    }

    char* p = (char*)GlobalLock(hMem);
    if (p == NULL) {
        GlobalFree(hMem);
        CloseClipboard();
        return false;
    }

    memcpy(p, text.c_str(), text.size() + 1);
    GlobalUnlock(hMem);
    SetClipboardData(CF_TEXT, hMem);
    CloseClipboard();
    return true;
}

bool CDwgPropsPanel::PasteTextToAcad(const std::string& text) {
    if (text.empty()) return false;

    CopyTextToClipboard(text);

    if (m_lastTextInputHwnd != NULL && IsWindow(m_lastTextInputHwnd)) {
        if (SendMessageA(m_lastTextInputHwnd, EM_REPLACESEL, TRUE, (LPARAM)text.c_str()) != 0) {
            return true;
        }
        if (SendMessageA(m_lastTextInputHwnd, WM_PASTE, 0, 0) != 0) {
            return true;
        }
    }

    PrintUtf8("\nПоле скопировано в буфер обмена. Вставьте в текст Ctrl+V.");
    return false;
}

void CDwgPropsPanel::UpdateInsertButtonState() {
    if (m_hBtnInsertField == NULL || m_hList == NULL) return;

    bool enabled = false;
    const int sel = (int)SendMessageA(m_hList, LB_GETCURSEL, 0, 0);
    if (sel >= 0 && sel < (int)m_items.size()) {
        enabled = m_items[sel].isCustom && (IsTextInputActive() || (m_lastTextInputHwnd != NULL && IsWindow(m_lastTextInputHwnd)));
    }

    EnableWindow(m_hBtnInsertField, enabled ? TRUE : FALSE);
}

LRESULT CALLBACK CDwgPropsPanel::WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    CDwgPropsPanel* self = (CDwgPropsPanel*)GetWindowLong(hWnd, GWL_USERDATA);

    if (msg == WM_NCCREATE) {
        CREATESTRUCT* cs = (CREATESTRUCT*)lParam;
        self = (CDwgPropsPanel*)cs->lpCreateParams;
        SetWindowLong(hWnd, GWL_USERDATA, (LONG)self);
        if (self != NULL) {
            self->m_hWnd = hWnd;
        }
    }

    if (self != NULL) {
        return self->HandleMessage(msg, wParam, lParam);
    }

    return DefWindowProc(hWnd, msg, wParam, lParam);
}

LRESULT CDwgPropsPanel::HandleMessage(UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_CREATE: {
            SetTimer(m_hWnd, 1, 250, NULL);

            m_hList = CreateWindowExA(
                WS_EX_CLIENTEDGE,
                "LISTBOX",
                "",
                WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOINTEGRALHEIGHT | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
                8,
                78,
                100,
                100,
                m_hWnd,
                NULL,
                GetModuleHandle(NULL),
                NULL);

            m_hEditSearch = CreateWindowExA(
                WS_EX_CLIENTEDGE,
                "EDIT",
                "",
                WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                8,
                48,
                468,
                20,
                m_hWnd,
                (HMENU)1001,
                GetModuleHandle(NULL),
                NULL);

            m_hComboGroup = CreateWindowExA(
                WS_EX_CLIENTEDGE,
                "COMBOBOX",
                "",
                WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
                8,
                24,
                468,
                200,
                m_hWnd,
                (HMENU)1004,
                GetModuleHandle(NULL),
                NULL);

            m_hLblTrackedXlsx = CreateWindowExA(
                0,
                "STATIC",
                "",
                WS_CHILD | WS_VISIBLE | SS_LEFT,
                8,
                4,
                460,
                18,
                m_hWnd,
                NULL,
                GetModuleHandle(NULL),
                NULL);

            const std::string btnLoadText = Utf8ToAcp("Загрузить параметры из XLSX");
            m_hBtnLoadXlsx = CreateWindowExA(
                0,
                "BUTTON",
                btnLoadText.c_str(),
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                8,
                8,
                220,
                28,
                m_hWnd,
                (HMENU)1002,
                GetModuleHandle(NULL),
                NULL);

            const std::string btnInsertText = Utf8ToAcp("Вставить выбранное пользовательское ПОЛЕ");
            m_hBtnInsertField = CreateWindowExA(
                0,
                "BUTTON",
                btnInsertText.c_str(),
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                236,
                8,
                240,
                28,
                m_hWnd,
                (HMENU)1003,
                GetModuleHandle(NULL),
                NULL);
            EnableWindow(m_hBtnInsertField, FALSE);

            LOGFONT lf;
            ZeroMemory(&lf, sizeof(lf));
            HFONT base = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
            GetObject(base, sizeof(lf), &lf);

            lf.lfWeight = FW_BOLD;
            lf.lfHeight = lf.lfHeight - 2;
            m_fontNameBold = CreateFontIndirect(&lf);

            lf.lfWeight = FW_NORMAL;
            lf.lfHeight = lf.lfHeight - 1;
            m_fontDesc = CreateFontIndirect(&lf);

            lf.lfWeight = FW_NORMAL;
            lf.lfHeight = lf.lfHeight + 3;
            m_fontValue = CreateFontIndirect(&lf);

            if (m_hLblTrackedXlsx != NULL) {
                SendMessageA(m_hLblTrackedXlsx, WM_SETFONT, (WPARAM)(m_fontDesc != NULL ? m_fontDesc : GetStockObject(DEFAULT_GUI_FONT)), TRUE);
            }

            RebuildGroupFilter();
            RefreshTrackedXlsxState(true);
            return 0;
        }

        case WM_SIZE:
            if (m_hList != NULL) {
                const int cx = LOWORD(lParam);
                const int cy = HIWORD(lParam);
                const int buttonH = 28;
                const int gap = 8;
                MoveWindow(m_hList, 8, 74, cx - 16, cy - (buttonH + gap * 3 + 66), TRUE);

                if (m_hEditSearch != NULL) {
                    MoveWindow(m_hEditSearch, 8, 48, cx - 16, 20, TRUE);
                }

                if (m_hComboGroup != NULL) {
                    MoveWindow(m_hComboGroup, 8, 24, cx - 16, 200, TRUE);
                }

                if (m_hLblTrackedXlsx != NULL) {
                    MoveWindow(m_hLblTrackedXlsx, 8, 4, cx - 16, 18, TRUE);
                }

                if (m_hBtnLoadXlsx != NULL) {
                    MoveWindow(m_hBtnLoadXlsx, 8, cy - (buttonH + gap), 220, buttonH, TRUE);
                }
                if (m_hBtnInsertField != NULL) {
                    MoveWindow(m_hBtnInsertField, 236, cy - (buttonH + gap), 240, buttonH, TRUE);
                }
            }
            if (wParam != SIZE_MINIMIZED) {
                SaveCurrentBounds();
            }
            return 0;

        case WM_MOVE:
            SaveCurrentBounds();
            return 0;

        case WM_MEASUREITEM:
            if (wParam == 0) {
                LPMEASUREITEMSTRUCT mi = (LPMEASUREITEMSTRUCT)lParam;
                mi->itemHeight = 36;
                return TRUE;
            }
            return FALSE;

        case WM_DRAWITEM:
            if (wParam == 0) {
                DrawListItem((const DRAWITEMSTRUCT*)lParam);
                return TRUE;
            }
            return FALSE;

        case WM_COMMAND:
            if (LOWORD(wParam) == 1002 && HIWORD(wParam) == BN_CLICKED) {
                Xlsx2DwgProp_Command();
                ReloadProperties();
                RefreshTrackedXlsxState(true);
                Show();
            }
            if (LOWORD(wParam) == 1001 && HIWORD(wParam) == EN_CHANGE) {
                RebuildVisibleList();
                UpdateInsertButtonState();
            }
            if (LOWORD(wParam) == 1004 && HIWORD(wParam) == CBN_SELCHANGE) {
                RebuildVisibleList();
                UpdateInsertButtonState();
            }
            if (LOWORD(wParam) == 1003 && HIWORD(wParam) == BN_CLICKED) {
                const int sel = (int)SendMessageA(m_hList, LB_GETCURSEL, 0, 0);
                if (sel >= 0 && sel < (int)m_items.size() && m_items[sel].isCustom) {
                    std::string fieldCode = "%<\\AcVar CustomDP." + m_items[sel].key + " \\f \"%tc1\">%";
                    PasteTextToAcad(fieldCode);
                }
            }
            if (LOWORD(wParam) == 0 && HIWORD(wParam) == LBN_SELCHANGE) {
                UpdateInsertButtonState();
                RefreshTrackedXlsxState(false);
            }
            return 0;

        case WM_CTLCOLORSTATIC:
            if ((HWND)lParam == m_hLblTrackedXlsx) {
                HDC hdc = (HDC)wParam;
                SetBkMode(hdc, TRANSPARENT);
                SetTextColor(hdc, m_trackedXlsxHashMismatch ? RGB(220, 0, 0) : RGB(0, 0, 0));
                return (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
            }
            break;

        case WM_TIMER:
            if (wParam == 1) {
                UpdateTopMostState();

                GUITHREADINFO gi;
                ZeroMemory(&gi, sizeof(gi));
                gi.cbSize = sizeof(gi);
                if (GetGUIThreadInfo(0, &gi)) {
                    HWND f = gi.hwndFocus;
                    if (f != NULL && f != m_hWnd && f != m_hList && f != m_hBtnLoadXlsx && f != m_hBtnInsertField) {
                        m_lastTextInputHwnd = f;
                    }
                }

                UpdateInsertButtonState();
            }
            return 0;

        case WM_CLOSE:
            SaveCurrentBounds();
            ShowWindow(m_hWnd, SW_HIDE);
            return 0;

        case WM_DESTROY:
            KillTimer(m_hWnd, 1);
            return 0;
    }

    return DefWindowProc(m_hWnd, msg, wParam, lParam);
}
