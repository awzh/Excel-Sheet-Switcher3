// ExcelCom.h
#pragma once
#include <windows.h>
#include <comdef.h>
#include <atlbase.h>

// 导入 Excel 类型库
// {00020813-0000-0000-C000-000000000046} 是 Excel 的 CLSID
#import "libid:{00020813-0000-0000-C000-000000000046}" \
    auto_rename auto_search no_dual_interfaces raw_interfaces_only \
    rename("RGB", "ExcelRGB") rename("DialogBox", "ExcelDialogBox")

class ExcelCom
{
public:
    ExcelCom();
    ~ExcelCom();
    bool ConnectToRunningExcel();           // 连接到当前运行中的 Excel
    bool SwitchToPreviousSheet();           // 切换到上一个工作表
    void RecordCurrentSheet();              // 记录当前活动的工作表
    bool IsConnected() const { return m_pApp != nullptr; }

private:
    Excel::_ApplicationPtr m_pApp;
    Excel::_WorkbookPtr    m_pPrevWorkbook;
    Excel::_WorksheetPtr   m_pPrevSheet;
    bool                   m_bHasPrevious;
};
