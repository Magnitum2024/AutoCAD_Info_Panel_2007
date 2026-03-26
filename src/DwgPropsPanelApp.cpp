#include "DwgPropsPanel.h"
#include "ToolbarSetup.h"
#include "Xlsx2DwgProp.h"
#include <aced.h>
#include <rxregsvc.h>
#include <locale.h>
#include <tchar.h>



namespace {

const ACHAR* kCmdGroup = ACRX_T("DWG_PROPS_PANEL_CMDS");
CDwgPropsPanel g_panel;

void InitLocaleForCyrillic() {
    _tsetlocale(LC_ALL, _T("russian"));
}

void CmdShow() {
    g_panel.Create();
    g_panel.ReloadProperties();
    g_panel.Show();
}

void CmdHide() {
    g_panel.Hide();
}

void CmdImportXlsx() {
    Xlsx2DwgProp_Command();
    g_panel.ReloadProperties();
    g_panel.Show();
}

}  // namespace

void DwgPropsPanel_Show() { CmdShow(); }
void DwgPropsPanel_Hide() { CmdHide(); }

void DwgPropsPanel_Init(void* pkt) {
    InitLocaleForCyrillic();

    acrxUnlockApplication(pkt);
    acrxRegisterAppMDIAware(pkt);
    acedRegCmds->addCommand(kCmdGroup, ACRX_T("SHOWDWGPROPS"), ACRX_T("SHOWDWGPROPS"), ACRX_CMD_MODAL, CmdShow);
    acedRegCmds->addCommand(kCmdGroup, ACRX_T("HIDEDWGPROPS"), ACRX_T("HIDEDWGPROPS"), ACRX_CMD_MODAL, CmdHide);
    acedRegCmds->addCommand(kCmdGroup, ACRX_T("XLSX2DWGPROP"), ACRX_T("XLSX2DWGPROP"), ACRX_CMD_MODAL, CmdImportXlsx);

    EnsureMgPanelToolbar();
}

void DwgPropsPanel_Unload() {
    acedRegCmds->removeGroup(kCmdGroup);
    g_panel.Destroy();
}
