// ExcelCom.cpp
#include "ExcelCom.h"

ExcelCom::ExcelCom() : m_bHasPrevious(false)
{
    // 初始化 COM 环境（必须在调用 COM 前执行）
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
}

ExcelCom::~ExcelCom()
{
    // 清理资源
    m_pPrevSheet = nullptr;
    m_pPrevWorkbook = nullptr;
    m_pApp = nullptr;
    CoUninitialize();
}

bool ExcelCom::ConnectToRunningExcel()
{
    try
    {
        // CLSID 为 Excel.Application 的类 ID
        CLSID clsid;
        HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
        if (FAILED(hr)) return false;

        // 获取正在运行的 Excel 实例
        IUnknown* pUnk = nullptr;
        hr = GetActiveObject(clsid, NULL, &pUnk);
        if (FAILED(hr)) return false;

        // 转换为 Excel 应用程序对象
        hr = pUnk->QueryInterface(__uuidof(Excel::_Application), (void**)&m_pApp);
        pUnk->Release();

        if (FAILED(hr) || !m_pApp) return false;

        return true;
    }
    catch (_com_error&)
    {
        return false;
    }
}

void ExcelCom::RecordCurrentSheet()
{
    if (!m_pApp) return;

    try
    {
        m_pPrevWorkbook = m_pApp->ActiveWorkbook;
        m_pPrevSheet = m_pApp->ActiveSheet;
        m_bHasPrevious = true;
    }
    catch (_com_error&)
    {
        // 忽略错误（比如没有活动工作表）
    }
}

bool ExcelCom::SwitchToPreviousSheet()
{
    if (!m_pApp || !m_bHasPrevious) return false;

    try
    {
        // 保存当前工作表
        Excel::_WorkbookPtr curWorkbook = m_pApp->ActiveWorkbook;
        Excel::_WorksheetPtr curSheet = m_pApp->ActiveSheet;

        // 激活上一个工作表
        if (m_pPrevWorkbook)
            m_pPrevWorkbook->Activate();
        if (m_pPrevSheet)
            m_pPrevSheet->Activate();

        // 更新"上一个"为刚才保存的当前工作表
        m_pPrevWorkbook = curWorkbook;
        m_pPrevSheet = curSheet;
        return true;
    }
    catch (_com_error&)
    {
        return false;
    }
}
