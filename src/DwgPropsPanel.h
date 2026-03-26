#pragma once

#include <windows.h>
#include <string>
#include <vector>

class CDwgPropsPanel {
public:
    CDwgPropsPanel();
    ~CDwgPropsPanel();

    bool Create();
    void Destroy();
    void Show();
    void Hide();
    void ReloadProperties();

private:
    static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
    LRESULT HandleMessage(UINT msg, WPARAM wParam, LPARAM lParam);

    void AddProperty(const char* name, const std::string& value, bool isCustom, const std::string& description = "", const std::string& group = "");
    std::string ReadSysVarAsString(const char* varName) const;

    bool IsHostAppForeground() const;
    void UpdateTopMostState();
    void DrawListItem(const DRAWITEMSTRUCT* dis);
    bool IsTextInputActive() const;
    void UpdateInsertButtonState();
    bool PasteTextToAcad(const std::string& text); // direct insert into editor; if failed, keep text in clipboard for Ctrl+V
    bool CopyTextToClipboard(const std::string& text);
    bool LoadSavedBounds(int& x, int& y, int& w, int& h) const;
    void SaveCurrentBounds() const;
    void RefreshTrackedXlsxState(bool force);
    void RebuildVisibleList();
    void RebuildGroupFilter();

private:
    HWND m_hWnd;
    HWND m_hList;
    struct ListItem {
        std::string key;
        std::string value;
        bool isCustom;
        std::string description;
        std::string group;
    };

    HWND m_hBtnLoadXlsx;
    HWND m_hBtnInsertField;
    HWND m_hLblTrackedXlsx;
    HWND m_hEditSearch;
    HWND m_hComboGroup;
    HFONT m_fontNameBold;
    HFONT m_fontDesc;
    HFONT m_fontValue;
    bool m_isTopMost;
    bool m_trackedXlsxHashMismatch;
    DWORD m_lastTrackedXlsxCheckTick;
    std::vector<ListItem> m_allItems;
    std::vector<ListItem> m_items;
    HWND m_lastTextInputHwnd;
};

// Для встраивания в существующий acrxEntryPoint проекта.
void DwgPropsPanel_Init(void* pkt);
void DwgPropsPanel_Unload();
void DwgPropsPanel_Show();
void DwgPropsPanel_Hide();
