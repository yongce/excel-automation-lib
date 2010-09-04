/*!
* @file    ExcelWorkbook.cpp
* @brief   Implementation file for class ExcelWorkbook
* @date    2009-12-01
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#include <tchar.h>
#include <cassert>
#include <vector>

#include "ExcelWorkbook.h"
#include "ExcelWorksheetSet.h"
#include "ExcelWorksheet.h"
#include "ComUtil.h"
#include "Noncopyable.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


////////////////////////////////////////////////////////////////////////////////
// Definition and implementation of class ExcelWorkbookImpl

/*!
* @brief Class ExcelWorkbookImpl inplements ExcelWorkbook's interfaces.
*/
class ExcelWorkbookImpl : public BodyBase, public Noncopyable
{
    // All members are private. Only the friend class ExcelWorkbook can access members of ExcelWorkbookImpl.
    friend class ExcelWorkbook;

private:
    ExcelWorkbookImpl(IDispatch *pWorkbook): m_pWorkbook(pWorkbook)
    {
        assert(pWorkbook);
    }

    virtual ~ExcelWorkbookImpl()
    {
        if (m_pWorkbook)
        {
            m_pWorkbook->Release();
            m_pWorkbook = 0;
        }
    }

    ExcelWorksheet GetActiveWorksheet();

    ExcelWorksheetSet GetAllWorksheets();

    bool Save();
    bool SaveAs(const ELstring &filename);

    bool Close();


private:
    IDispatch *m_pWorkbook;
};


ExcelWorksheet ExcelWorkbookImpl::GetActiveWorksheet()
{
    assert(m_pWorkbook);

    VARIANT result;
    VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pWorkbook, DISPATCH_PROPERTYGET, OLESTR("ActiveSheet"), &result, 0);

    if (FAILED(hr))
        return ExcelWorksheet();

    return ExcelWorksheet(result.pdispVal);
}


ExcelWorksheetSet ExcelWorkbookImpl::GetAllWorksheets()
{
    assert(m_pWorkbook);

    VARIANT result;
    VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pWorkbook, DISPATCH_PROPERTYGET, OLESTR("Worksheets"), &result, 0);

    if (FAILED(hr))
        return ExcelWorksheetSet();

    return ExcelWorksheetSet(result.pdispVal);
}


bool ExcelWorkbookImpl::Save()
{
    assert(m_pWorkbook);

    HRESULT hr = ComUtil::Invoke(m_pWorkbook, DISPATCH_METHOD, OLESTR("Save"), NULL, 0);

    return SUCCEEDED(hr);
}


bool ExcelWorkbookImpl::SaveAs(const ELstring &filename)
{
    assert(m_pWorkbook);

    // Get full path name for the file
    std::vector<ELchar> fullpath(256, ELtext('\0'));
    DWORD bufLen = ::GetFullPathName(filename.c_str(), fullpath.size(), &fullpath[0], 0);
    if (bufLen == 0)
        return false;  // the file cannot be found

    if (bufLen > fullpath.size())
    {
        fullpath.resize(bufLen, ELtext('\0'));
        ::GetFullPathName(filename.c_str(), fullpath.size(), &fullpath[0], 0);
    }

    VARIANT param;
    param.vt = VT_BSTR;
    param.bstrVal = ::SysAllocString(&fullpath[0]);

    HRESULT hr = ComUtil::Invoke(m_pWorkbook, DISPATCH_METHOD, OLESTR("SaveAs"), NULL, 1, param);

    ::VariantClear(&param);

    return SUCCEEDED(hr);
}


bool ExcelWorkbookImpl::Close()
{
    assert(m_pWorkbook);

    HRESULT hr = ComUtil::Invoke(m_pWorkbook, DISPATCH_METHOD, OLESTR("Close"), NULL, 0);

    m_pWorkbook->Release();
    m_pWorkbook = 0;

    return SUCCEEDED(hr);
}


////////////////////////////////////////////////////////////////////////////////
// Implementation of class ExcelWorkbook

ExcelWorkbook::ExcelWorkbook(IDispatch *pWorkbook): HandleBase(new ExcelWorkbookImpl(pWorkbook))
{
    assert(pWorkbook);
}


ExcelWorksheet ExcelWorkbook::GetActiveWorksheet() const
{
    return Body().GetActiveWorksheet();
}


ExcelWorksheetSet ExcelWorkbook::GetAllWorksheets() const
{
    return Body().GetAllWorksheets();
}


bool ExcelWorkbook::Save() const
{
    return Body().Save();
}


bool ExcelWorkbook::SaveAs(const ELstring &filename)
{
    return Body().SaveAs(filename);
}


bool ExcelWorkbook::Close() const
{
    return Body().Close();
}


// <begin> Handle/Body pattern implementation

ExcelWorkbook::ExcelWorkbook(ExcelWorkbookImpl *impl): HandleBase(impl)
{ 
}


ExcelWorkbookImpl& ExcelWorkbook::Body() const
{
    return dynamic_cast<ExcelWorkbookImpl&>(HandleBase::Body());
}

// <end> Handle/Body pattern implementation


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END
