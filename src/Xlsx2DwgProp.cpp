#include "Xlsx2DwgProp.h"

#include <aced.h>
#include <adslib.h>
#include <dbapserv.h>
#include <summinfo.h>

#include <ole2.h>
#include <oleauto.h>
#include <commdlg.h>

#include <string>
#include <map>
#include <vector>
#include <fstream>
#include <stdlib.h>
#include <stdarg.h>

namespace {


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

void PrintUtf8(const char* fmtUtf8, ...) {
    std::string fmt = Utf8ToAcp(fmtUtf8);

    char out[2048];
    out[0] = '\0';

    va_list args;
    va_start(args, fmtUtf8);
    _vsnprintf(out, sizeof(out) - 1, fmt.c_str(), args);
    va_end(args);

    out[sizeof(out) - 1] = '\0';
#ifdef AD_UNICODE
    int wlen = MultiByteToWideChar(CP_ACP, 0, out, -1, NULL, 0);
    if (wlen > 0) {
        std::wstring outW;
        outW.resize(wlen - 1);
        MultiByteToWideChar(CP_ACP, 0, out, -1, &outW[0], wlen);
        acutPrintf(ACRX_T("%s"), outW.c_str());
    } else {
        acutPrintf(ACRX_T("%s"), ACRX_T(""));
    }
#else
    acutPrintf("%s", out);
#endif
}

struct PropPair {
    std::string key;
    std::string value;
    std::string description;
};

std::map<std::string, std::string> g_lastDescriptions;
std::map<std::string, std::string> g_lastGroups;
std::string g_lastReadError;

struct ImportSettings {
    std::string worksheetName;
    long keyColumn;
    long valueColumn;
    long worksheetIndex;
    std::string reader;
    std::string libreOfficePath;
    std::string configPath;
};

std::string HResultHex(HRESULT hr) {
    char buf[32];
    _snprintf(buf, sizeof(buf) - 1, "0x%08lX", (unsigned long)hr);
    buf[sizeof(buf) - 1] = '\0';
    return std::string(buf);
}

std::string ModuleDir() {
    char path[MAX_PATH];
    path[0] = '\0';
    GetModuleFileName(NULL, path, MAX_PATH);

    std::string p(path);
    const std::string::size_type pos = p.find_last_of("\\/");
    if (pos == std::string::npos) {
        return std::string(".");
    }
    return p.substr(0, pos);
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

std::string PanelIniPath() {
    return UserConfigDir() + "\\DwgPropsPanel.ini";
}

std::string StripUtf8Bom(const std::string& s) {
    if (s.size() >= 3 &&
        (unsigned char)s[0] == 0xEF &&
        (unsigned char)s[1] == 0xBB &&
        (unsigned char)s[2] == 0xBF) {
        return s.substr(3);
    }
    return s;
}

std::string FileNameOnly(const std::string& fullPath) {
    const std::string::size_type p = fullPath.find_last_of("\\/");
    if (p == std::string::npos) return fullPath;
    return fullPath.substr(p + 1);
}

std::string FileDirOnly(const std::string& fullPath) {
    const std::string::size_type p = fullPath.find_last_of("\\/");
    if (p == std::string::npos) return std::string(".");
    return fullPath.substr(0, p);
}

unsigned long Crc32File(const std::string& path, bool& ok) {
    ok = false;
    std::ifstream in(path.c_str(), std::ios::binary);
    if (!in) return 0;

    unsigned long crc = 0xFFFFFFFFUL;
    char buf[4096];
    while (in) {
        in.read(buf, sizeof(buf));
        std::streamsize n = in.gcount();
        for (std::streamsize i = 0; i < n; ++i) {
            unsigned long byte = (unsigned char)buf[i];
            crc ^= byte;
            for (int k = 0; k < 8; ++k) {
                unsigned long mask = (unsigned long)-(long)(crc & 1UL);
                crc = (crc >> 1) ^ (0xEDB88320UL & mask);
            }
        }
    }

    ok = true;
    return ~crc;
}

bool RunProcessAndWait(const std::string& exePath, const std::string& args) {
    STARTUPINFOA si;
    PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof(si));
    ZeroMemory(&pi, sizeof(pi));
    si.cb = sizeof(si);

    std::string cmd = "\"" + exePath + "\" " + args;
    std::vector<char> cmdBuf(cmd.begin(), cmd.end());
    cmdBuf.push_back('\0');

    if (!CreateProcessA(NULL, &cmdBuf[0], NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi)) {
        return false;
    }

    WaitForSingleObject(pi.hProcess, INFINITE);
    DWORD code = 1;
    GetExitCodeProcess(pi.hProcess, &code);
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return code == 0;
}

bool TryLibreOfficeConvert(const std::string& fileName, const std::string& outDir, const std::string& configuredPath) {
    const std::string args = "--headless --convert-to csv --outdir \"" + outDir + "\" \"" + fileName + "\"";
    std::vector<std::string> candidates;
    if (!configuredPath.empty()) {
        candidates.push_back(configuredPath);
        std::string configuredUpper = configuredPath;
        for (size_t i = 0; i < configuredUpper.size(); ++i) {
            if (configuredUpper[i] >= 'a' && configuredUpper[i] <= 'z') {
                configuredUpper[i] = (char)(configuredUpper[i] - ('a' - 'A'));
            }
        }
        const std::string scalcExe = "SCALC.EXE";
        if (configuredUpper.size() >= scalcExe.size() &&
            configuredUpper.substr(configuredUpper.size() - scalcExe.size()) == scalcExe) {
            std::string sofficePath = configuredPath;
            sofficePath.replace(sofficePath.size() - 9, 9, "soffice.exe");
            candidates.push_back(sofficePath);
        }
    }
    if (configuredPath == "soffice" || configuredPath == "soffice.exe" || configuredPath == "scalc.exe" || configuredPath.empty()) {
        candidates.push_back("soffice.exe");
        candidates.push_back("soffice");
        candidates.push_back("scalc.exe");
        candidates.push_back("C:\\Program Files\\LibreOffice\\program\\soffice.exe");
        candidates.push_back("C:\\Program Files\\LibreOffice\\program\\scalc.exe");
        candidates.push_back("C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe");
        candidates.push_back("C:\\Program Files (x86)\\LibreOffice\\program\\scalc.exe");
    }

    for (size_t i = 0; i < candidates.size(); ++i) {
        if (RunProcessAndWait(candidates[i], args)) {
            return true;
        }
    }
    return false;
}

ImportSettings LoadSettings() {
    ImportSettings st;
    st.worksheetName = "";
    st.keyColumn = 2;
    st.valueColumn = 5;
    st.worksheetIndex = 1;
    st.reader = "Excel";
    st.libreOfficePath = "C:\\Program Files\\LibreOffice\\program\\soffice.exe";

    st.configPath = ModuleDir() + "\\Xlsx2DwgProp.ini";

    char buf[256];
    buf[0] = '\0';
    GetPrivateProfileString("XLSX", "WorksheetName", st.worksheetName.c_str(), buf, sizeof(buf), st.configPath.c_str());
    if (buf[0] != '\0') {
        st.worksheetName = buf;
    }

    st.keyColumn = GetPrivateProfileInt("XLSX", "KeyColumn", (int)st.keyColumn, st.configPath.c_str());
    st.valueColumn = GetPrivateProfileInt("XLSX", "ValueColumn", (int)st.valueColumn, st.configPath.c_str());
    st.worksheetIndex = GetPrivateProfileInt("XLSX", "WorksheetIndex", (int)st.worksheetIndex, st.configPath.c_str());
    GetPrivateProfileString("XLSX", "Reader", st.reader.c_str(), buf, sizeof(buf), st.configPath.c_str());
    if (buf[0] != '\0') st.reader = buf;
    GetPrivateProfileString("XLSX", "LibreOfficePath", st.libreOfficePath.c_str(), buf, sizeof(buf), st.configPath.c_str());
    if (buf[0] != '\0') st.libreOfficePath = buf;

    if (st.keyColumn <= 0) st.keyColumn = 2;
    if (st.valueColumn <= 0) st.valueColumn = 5;
    if (st.worksheetIndex <= 0) st.worksheetIndex = 1;

    return st;
}

bool IsStandardProp(const std::string& keyUpper) {
    return keyUpper == "TITLE" || keyUpper == "SUBJECT" || keyUpper == "AUTHOR" ||
           keyUpper == "KEYWORDS" || keyUpper == "COMMENTS" || keyUpper == "REVISIONNUMBER" ||
           keyUpper == "HYPERLINKBASE";
}

std::string ToUpper(const std::string& s) {
    std::string out = s;
    for (size_t i = 0; i < out.size(); ++i) {
        if (out[i] >= 'a' && out[i] <= 'z') {
            out[i] = (char)(out[i] - ('a' - 'A'));
        }
    }
    return out;
}

std::string Trim(const std::string& s) {
    size_t b = 0;
    while (b < s.size() && (s[b] == ' ' || s[b] == '\t' || s[b] == '\r' || s[b] == '\n')) {
        ++b;
    }
    size_t e = s.size();
    while (e > b && (s[e - 1] == ' ' || s[e - 1] == '\t' || s[e - 1] == '\r' || s[e - 1] == '\n')) {
        --e;
    }
    return s.substr(b, e - b);
}

std::wstring Utf16(const std::string& s) {
    if (s.empty()) {
        return std::wstring();
    }
    const int n = MultiByteToWideChar(CP_ACP, 0, s.c_str(), -1, NULL, 0);
    std::wstring ws;
    ws.resize(n > 0 ? n - 1 : 0);
    if (n > 0) {
        MultiByteToWideChar(CP_ACP, 0, s.c_str(), -1, &ws[0], n);
    }
    return ws;
}


std::basic_string<ACHAR> ToAChar(const std::string& s) {
#ifdef AD_UNICODE
    std::wstring w = Utf16(s);
    return std::basic_string<ACHAR>(w.c_str());
#else
    return std::basic_string<ACHAR>(s.c_str());
#endif
}

std::string VariantToString(const VARIANT& v) {
    VARIANT tmp;
    VariantInit(&tmp);

    if (FAILED(VariantChangeType(&tmp, const_cast<VARIANT*>(&v), 0, VT_BSTR))) {
        return "";
    }

    std::string out;
    if (tmp.vt == VT_BSTR && tmp.bstrVal != NULL) {
        const int n = WideCharToMultiByte(CP_ACP, 0, tmp.bstrVal, -1, NULL, 0, NULL, NULL);
        if (n > 0) {
            out.resize(n - 1);
            WideCharToMultiByte(CP_ACP, 0, tmp.bstrVal, -1, &out[0], n, NULL, NULL);
        }
    }

    VariantClear(&tmp);
    return Trim(out);
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

bool ReadFromXlsx(const std::string& fileName, const ImportSettings& settings, std::vector<PropPair>& outProps) {
    g_lastReadError.clear();
    const std::string reader = ToUpper(Trim(settings.reader));
    if (reader == "LIBREOFFICE") {
        // LibreOffice mode: convert XLSX to CSV in temp folder, then parse columns.
        char tmpDirBuf[MAX_PATH];
        tmpDirBuf[0] = '\0';
        GetTempPathA(MAX_PATH, tmpDirBuf);
        std::string tmpDir(tmpDirBuf);
        tmpDir += "Xlsx2DwgPropTmp";
        CreateDirectoryA(tmpDir.c_str(), NULL);

        if (!TryLibreOfficeConvert(fileName, tmpDir, settings.libreOfficePath)) {
            g_lastReadError = Utf8ToAcp("Ошибка конвертации LibreOffice. Проверьте Reader=LibreOffice и LibreOfficePath.");
            return false;
        }

        std::string csvPath = tmpDir + "\\";
        csvPath += FileNameOnly(fileName);
        const std::string::size_type p = csvPath.find_last_of('.');
        if (p != std::string::npos) {
            csvPath = csvPath.substr(0, p);
        }
        csvPath += ".csv";

        FILE* f = fopen(csvPath.c_str(), "rb");
        if (f == NULL) {
            g_lastReadError = Utf8ToAcp("LibreOffice не создал CSV-файл.");
            return false;
        }

        g_lastDescriptions.clear();
        g_lastGroups.clear();

        char line[8192];
        bool firstLine = true;
        while (fgets(line, sizeof(line), f) != NULL) {
            std::vector<std::string> cols;
            std::string cur;
            bool inQ = false;
            for (size_t i = 0; line[i] != '\0'; ++i) {
                const char ch = line[i];
                if (ch == '\"') {
                    inQ = !inQ;
                    continue;
                }
                if ((ch == ';' || ch == ',') && !inQ) {
                    cols.push_back(cur);
                    cur.clear();
                    continue;
                }
                if (ch == '\r' || ch == '\n') {
                    continue;
                }
                cur.push_back(ch);
            }
            cols.push_back(cur);
            if (firstLine && !cols.empty()) {
                cols[0] = StripUtf8Bom(cols[0]);
            }
            firstLine = false;

            const int k = (int)settings.keyColumn - 1;
            const int v = (int)settings.valueColumn - 1;
            const int g = 0; // column A
            const int d = 2; // column C
            if (k < 0 || v < 0 || k >= (int)cols.size() || v >= (int)cols.size()) {
                continue;
            }

            std::string key = Utf8ToAcp(Trim(cols[k]).c_str());
            std::string val = Utf8ToAcp(Trim(cols[v]).c_str());
            std::string group = (g < (int)cols.size()) ? Utf8ToAcp(Trim(cols[g]).c_str()) : "";
            std::string descr = (d < (int)cols.size()) ? Utf8ToAcp(Trim(cols[d]).c_str()) : "";

            if (!key.empty() && !val.empty()) {
                PropPair p2;
                p2.key = key;
                p2.value = val;
                p2.description = descr;
                outProps.push_back(p2);
                g_lastDescriptions[ToUpper(key)] = descr;
                g_lastGroups[ToUpper(key)] = group;
            }
        }
        fclose(f);
        if (outProps.empty()) {
            g_lastReadError = Utf8ToAcp("Чтение через LibreOffice завершено, но строки данных не найдены.");
        }
        return true;
    }

    g_lastDescriptions.clear();
    g_lastGroups.clear();

    HRESULT hr = CoInitialize(NULL);
    const bool mustUninit = SUCCEEDED(hr);

    CLSID clsid;
    hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    if (FAILED(hr)) {
        g_lastReadError = Utf8ToAcp("Excel COM недоступен") + " (CLSIDFromProgID=" + HResultHex(hr) + ").";
        if (mustUninit) CoUninitialize();
        return false;
    }

    IDispatch* excel = NULL;
    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&excel);
    if (FAILED(hr) || excel == NULL) {
        g_lastReadError = Utf8ToAcp("Не удалось создать COM-экземпляр Excel") + " (CoCreateInstance=" + HResultHex(hr) + ").";
        if (mustUninit) CoUninitialize();
        return false;
    }

    VARIANT x;
    VariantInit(&x);
    x.vt = VT_BOOL;
    x.boolVal = VARIANT_FALSE;
    AutoWrap(DISPATCH_PROPERTYPUT, NULL, excel, L"Visible", 1, x);

    VARIANT vWorkbooks;
    VariantInit(&vWorkbooks);
    hr = AutoWrap(DISPATCH_PROPERTYGET, &vWorkbooks, excel, L"Workbooks", 0);
    if (FAILED(hr)) {
        g_lastReadError = Utf8ToAcp("Не удалось получить коллекцию Workbooks у Excel") + " (Workbooks get=" + HResultHex(hr) + ").";
        excel->Release();
        if (mustUninit) CoUninitialize();
        return false;
    }

    IDispatch* workbooks = vWorkbooks.pdispVal;
    VARIANT vFile;
    VariantInit(&vFile);
    vFile.vt = VT_BSTR;
    std::wstring wfn = Utf16(fileName);
    vFile.bstrVal = SysAllocString(wfn.c_str());

    VARIANT vWorkbook;
    VariantInit(&vWorkbook);
    hr = AutoWrap(DISPATCH_METHOD, &vWorkbook, workbooks, L"Open", 1, vFile);
    if (FAILED(hr)) {
        g_lastReadError = Utf8ToAcp("Excel не смог открыть XLSX-файл") + " (Open=" + HResultHex(hr) + ").";
        SysFreeString(vFile.bstrVal);
        workbooks->Release();
        excel->Release();
        if (mustUninit) CoUninitialize();
        return false;
    }
    SysFreeString(vFile.bstrVal);

    IDispatch* workbook = vWorkbook.pdispVal;

    VARIANT vWorksheets;
    VariantInit(&vWorksheets);
    AutoWrap(DISPATCH_PROPERTYGET, &vWorksheets, workbook, L"Worksheets", 0);
    IDispatch* worksheets = vWorksheets.pdispVal;

    VARIANT vSheet;
    VariantInit(&vSheet);

    if (!settings.worksheetName.empty()) {
        VARIANT vSheetName;
        VariantInit(&vSheetName);
        vSheetName.vt = VT_BSTR;
        std::wstring wsName = Utf16(settings.worksheetName);
        vSheetName.bstrVal = SysAllocString(wsName.c_str());
        hr = AutoWrap(DISPATCH_PROPERTYGET, &vSheet, worksheets, L"Item", 1, vSheetName);
        SysFreeString(vSheetName.bstrVal);
    } else {
        VARIANT vSheetIndex;
        VariantInit(&vSheetIndex);
        vSheetIndex.vt = VT_I4;
        vSheetIndex.lVal = settings.worksheetIndex;
        hr = AutoWrap(DISPATCH_PROPERTYGET, &vSheet, worksheets, L"Item", 1, vSheetIndex);
    }

    if (FAILED(hr)) {
        g_lastReadError = Utf8ToAcp("Лист не найден") + " (Item=" + HResultHex(hr) + ").";
        VARIANT vFalse;
        VariantInit(&vFalse);
        vFalse.vt = VT_BOOL;
        vFalse.boolVal = VARIANT_FALSE;
        AutoWrap(DISPATCH_METHOD, NULL, workbook, L"Close", 1, vFalse);
        workbook->Release();
        workbooks->Release();
        excel->Release();
        if (mustUninit) CoUninitialize();
        return false;
    }

    IDispatch* sheet = vSheet.pdispVal;

    VARIANT vUsed;
    VariantInit(&vUsed);
    AutoWrap(DISPATCH_PROPERTYGET, &vUsed, sheet, L"UsedRange", 0);
    IDispatch* used = vUsed.pdispVal;

    VARIANT vRows;
    VariantInit(&vRows);
    AutoWrap(DISPATCH_PROPERTYGET, &vRows, used, L"Rows", 0);
    IDispatch* rows = vRows.pdispVal;

    VARIANT vCount;
    VariantInit(&vCount);
    AutoWrap(DISPATCH_PROPERTYGET, &vCount, rows, L"Count", 0);
    const long rowCount = (vCount.vt == VT_I4) ? vCount.lVal : 0;

    for (long r = 1; r <= rowCount; ++r) {
        VARIANT vr, vc;
        VariantInit(&vr);
        VariantInit(&vc);

        vr.vt = VT_I4;
        vr.lVal = r;
        vc.vt = VT_I4;
        vc.lVal = settings.keyColumn;

        VARIANT vCell;
        VariantInit(&vCell);
        AutoWrap(DISPATCH_PROPERTYGET, &vCell, sheet, L"Cells", 2, vc, vr);
        IDispatch* cell = vCell.pdispVal;
        VARIANT vVal;
        VariantInit(&vVal);
        AutoWrap(DISPATCH_PROPERTYGET, &vVal, cell, L"Value", 0);
        std::string key = VariantToString(vVal);
        cell->Release();
        VariantClear(&vVal);

        vc.lVal = settings.valueColumn;
        VariantInit(&vCell);
        AutoWrap(DISPATCH_PROPERTYGET, &vCell, sheet, L"Cells", 2, vc, vr);
        cell = vCell.pdispVal;
        VariantInit(&vVal);
        AutoWrap(DISPATCH_PROPERTYGET, &vVal, cell, L"Value", 0);
        std::string val = VariantToString(vVal);
        cell->Release();
        VariantClear(&vVal);

        // Column A: parameter group name.
        vc.lVal = 1;
        VariantInit(&vCell);
        AutoWrap(DISPATCH_PROPERTYGET, &vCell, sheet, L"Cells", 2, vc, vr);
        cell = vCell.pdispVal;
        VariantInit(&vVal);
        AutoWrap(DISPATCH_PROPERTYGET, &vVal, cell, L"Value", 0);
        std::string group = VariantToString(vVal);
        cell->Release();
        VariantClear(&vVal);

        // Column C: attribute description.
        vc.lVal = 3;
        VariantInit(&vCell);
        AutoWrap(DISPATCH_PROPERTYGET, &vCell, sheet, L"Cells", 2, vc, vr);
        cell = vCell.pdispVal;
        VariantInit(&vVal);
        AutoWrap(DISPATCH_PROPERTYGET, &vVal, cell, L"Value", 0);
        std::string descr = VariantToString(vVal);
        cell->Release();
        VariantClear(&vVal);

        key = Trim(key);
        val = Trim(val);
        group = Trim(group);
        descr = Trim(descr);
        if (!key.empty() && !val.empty()) {
            PropPair p;
            p.key = key;
            p.value = val;
            p.description = descr;
            outProps.push_back(p);

            g_lastDescriptions[ToUpper(key)] = descr;
            g_lastGroups[ToUpper(key)] = group;
        }
    }

    VARIANT vFalse;
    VariantInit(&vFalse);
    vFalse.vt = VT_BOOL;
    vFalse.boolVal = VARIANT_FALSE;
    AutoWrap(DISPATCH_METHOD, NULL, workbook, L"Close", 1, vFalse);

    workbook->Release();
    workbooks->Release();
    excel->Release();

    if (mustUninit) CoUninitialize();
    return true;
}

Acad::ErrorStatus SetStandardProp(AcDbDatabaseSummaryInfo* pInfo, const std::string& keyUpper, const std::string& value) {
    const std::basic_string<ACHAR> w = ToAChar(value);
    const ACHAR* v = w.c_str();

    if (keyUpper == "TITLE") return pInfo->setTitle(v);
    if (keyUpper == "SUBJECT") return pInfo->setSubject(v);
    if (keyUpper == "AUTHOR") return pInfo->setAuthor(v);
    if (keyUpper == "KEYWORDS") return pInfo->setKeywords(v);
    if (keyUpper == "COMMENTS") return pInfo->setComments(v);
    if (keyUpper == "REVISIONNUMBER") return pInfo->setRevisionNumber(v);
    if (keyUpper == "HYPERLINKBASE") return pInfo->setHyperlinkBase(v);

    return Acad::eInvalidInput;
}

void ApplyProps(const std::vector<PropPair>& props) {
    AcDbDatabase* pDb = acdbHostApplicationServices()->workingDatabase();
    if (pDb == NULL) {
        PrintUtf8("\nНе удалось получить текущую базу DWG.");
        return;
    }

    AcDbDatabaseSummaryInfo* pInfo = NULL;
    if (acdbGetSummaryInfo(pDb, pInfo) != Acad::eOk || pInfo == NULL) {
        PrintUtf8("\nНе удалось получить SummaryInfo.");
        return;
    }

    int ok = 0;
    int err = 0;

    for (size_t i = 0; i < props.size(); ++i) {
        const std::string key = Trim(props[i].key);
        const std::string val = props[i].value;
        const std::string up = ToUpper(key);

        Acad::ErrorStatus es = Acad::eOk;

        if (IsStandardProp(up)) {
            es = SetStandardProp(pInfo, up, val);
        } else {
            std::basic_string<ACHAR> wk = ToAChar(key);
            std::basic_string<ACHAR> wv = ToAChar(val);
            es = pInfo->setCustomSummaryInfo(wk.c_str(), wv.c_str());
            if (es != Acad::eOk) {
                es = pInfo->addCustomSummaryInfo(wk.c_str(), wv.c_str());
            }
        }

        if (es == Acad::eOk) ++ok;
        else ++err;
    }

    const Acad::ErrorStatus putEs = acdbPutSummaryInfo(pInfo);
    delete pInfo;

    if (putEs != Acad::eOk) {
        PrintUtf8("\nНе удалось сохранить SummaryInfo.");
        return;
    }

    PrintUtf8("\nСвойств записано: %d, ошибок: %d.", ok, err);
    acedCommand(RTSTR, ACRX_T("_.REGEN"), RTNONE);
}

}  // namespace

void Xlsx2DwgProp_Command() {
    char fn[MAX_PATH];
    fn[0] = '\0';

    OPENFILENAMEA ofn;
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = GetForegroundWindow();
    ofn.lpstrFilter = "Файлы Excel (*.xlsx)\0*.xlsx\0Все файлы (*.*)\0*.*\0";
    ofn.lpstrFile = fn;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY;
    ofn.lpstrDefExt = "xlsx";
    std::string dlgTitle = Utf8ToAcp("Выберите XLSX файл");
    ofn.lpstrTitle = dlgTitle.c_str();

    if (!GetOpenFileNameA(&ofn)) {
        PrintUtf8("\nИмпорт из XLSX отменен.");
        return;
    }

    ImportSettings settings = LoadSettings();

    std::vector<PropPair> props;
    if (!ReadFromXlsx(fn, settings, props)) {
        const std::string excelErr = g_lastReadError;
        const std::string readerUpper = ToUpper(Trim(settings.reader));
        if (readerUpper != "LIBREOFFICE") {
            ImportSettings fallback = settings;
            fallback.reader = "LibreOffice";
            props.clear();
            if (ReadFromXlsx(fn, fallback, props)) {
                PrintUtf8("\nЧтение через Excel недоступно, выполнен переход на LibreOffice.");
                settings = fallback;
            } else {
                if (!excelErr.empty() && !g_lastReadError.empty()) {
                    g_lastReadError = excelErr + " Переход на LibreOffice тоже завершился ошибкой: " + g_lastReadError;
                } else if (!excelErr.empty()) {
                    g_lastReadError = excelErr;
                }
            }
        }
    }
    if (props.empty() && !g_lastReadError.empty()) {
        if (!g_lastReadError.empty()) {
            PrintUtf8("\nНе удалось прочитать XLSX: %s", g_lastReadError.c_str());
        } else {
            PrintUtf8("\nНе удалось прочитать XLSX. Проверьте Excel/LibreOffice и настройки листа в INI.");
        }
        PrintUtf8("\nДиагностика: Reader=%s, LibreOfficePath=%s, WorksheetIndex=%ld, WorksheetName=%s, KeyColumn=%ld, ValueColumn=%ld",
            settings.reader.c_str(),
            settings.libreOfficePath.c_str(),
            settings.worksheetIndex,
            settings.worksheetName.empty() ? "<пусто>" : settings.worksheetName.c_str(),
            settings.keyColumn,
            settings.valueColumn);
        PrintUtf8("\nINI: %s", settings.configPath.c_str());
        PrintUtf8("\nФайл: %s", fn);
        return;
    }

    if (props.empty()) {
        PrintUtf8("\nНа листе нет данных в настроенных колонках. КлючеваяКолонка=%ld, КолонкаЗначения=%ld",
            settings.keyColumn, settings.valueColumn);
        return;
    }

    ApplyProps(props);
    bool hashOk = false;
    const unsigned long crc = Crc32File(fn, hashOk);
    if (hashOk) {
        char num[64];
        _snprintf(num, sizeof(num) - 1, "%lu", crc);
        num[sizeof(num) - 1] = '\0';

        const std::string ini = PanelIniPath();
        WritePrivateProfileString("XLSX", "LastFileName", FileNameOnly(fn).c_str(), ini.c_str());
        WritePrivateProfileString("XLSX", "LastFileFullPath", fn, ini.c_str());
        WritePrivateProfileString("XLSX", "LastFileDir", FileDirOnly(fn).c_str(), ini.c_str());
        WritePrivateProfileString("XLSX", "LastFileCrc32", num, ini.c_str());
        if (GetPrivateProfileInt("XLSX", "HashCheckMinutes", 0, ini.c_str()) <= 0) {
            WritePrivateProfileString("XLSX", "HashCheckMinutes", "10", ini.c_str());
        }
    }
    PrintUtf8("\nИмпорт из XLSX завершен: %s", fn);
    PrintUtf8("\nИспользован конфиг: %s", settings.configPath.c_str());
}

bool Xlsx2DwgProp_GetDescriptionForKey(const std::string& key, std::string& outDescription) {
    outDescription.clear();
    const std::map<std::string, std::string>::const_iterator it = g_lastDescriptions.find(ToUpper(Trim(key)));
    if (it == g_lastDescriptions.end()) {
        return false;
    }
    outDescription = it->second;
    return !outDescription.empty();
}

bool Xlsx2DwgProp_GetTrackedFileStatus(std::string& fileNameOnly, std::string& fullPath, bool& hashMismatch) {
    hashMismatch = false;
    fileNameOnly.clear();
    fullPath.clear();

    const std::string ini = PanelIniPath();
    char pathBuf[MAX_PATH];
    pathBuf[0] = '\0';
    GetPrivateProfileString("XLSX", "LastFileFullPath", "", pathBuf, sizeof(pathBuf), ini.c_str());
    if (pathBuf[0] == '\0') {
        return false;
    }

    fullPath = pathBuf;
    fileNameOnly = FileNameOnly(fullPath);

    const unsigned long stored = (unsigned long)GetPrivateProfileInt("XLSX", "LastFileCrc32", 0, ini.c_str());
    bool ok = false;
    const unsigned long current = Crc32File(fullPath, ok);
    if (!ok) {
        hashMismatch = true;
        return true;
    }
    hashMismatch = stored != current;
    return true;
}

int Xlsx2DwgProp_GetHashCheckMinutes() {
    const std::string ini = PanelIniPath();
    int m = GetPrivateProfileInt("XLSX", "HashCheckMinutes", 10, ini.c_str());
    if (m <= 0) m = 10;
    return m;
}

bool Xlsx2DwgProp_GetGroupForKey(const std::string& key, std::string& outGroup) {
    outGroup.clear();
    const std::map<std::string, std::string>::const_iterator it = g_lastGroups.find(ToUpper(Trim(key)));
    if (it == g_lastGroups.end()) {
        return false;
    }
    outGroup = it->second;
    return !outGroup.empty();
}

void Xlsx2DwgProp_GetGroups(std::vector<std::string>& outGroups) {
    outGroups.clear();
    for (std::map<std::string, std::string>::const_iterator it = g_lastGroups.begin(); it != g_lastGroups.end(); ++it) {
        const std::string& g = it->second;
        if (g.empty()) {
            continue;
        }

        bool exists = false;
        for (size_t i = 0; i < outGroups.size(); ++i) {
            if (outGroups[i] == g) {
                exists = true;
                break;
            }
        }
        if (!exists) {
            outGroups.push_back(g);
        }
    }
}
