/*!
* @file    ExcelWorksheet.cpp
* @brief   Implementation file for class ExcelWorksheet and class ExcelWorksheetImpl
* @date    2009-12-01
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#include <tchar.h>
#include <cassert>

#include "ExcelWorksheet.h"
#include "ExcelRange.h"
#include "ExcelCell.h"
#include "ComUtil.h"
#include "Noncopyable.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


////////////////////////////////////////////////////////////////////////////////
// Definition and implementation of class ExcelWorksheetImpl

/*!
* @brief Class ExcelWorksheetImpl inplements ExcelWorksheet's interfaces.
*/
class ExcelWorksheetImpl : public BodyBase, public Noncopyable
{
    // All members are private. Only the friend class ExcelWorksheet can access members of ExcelWorksheetImpl.
    friend class ExcelWorksheet;

private:
    ExcelWorksheetImpl(IDispatch *pWorksheet): m_pWorksheet(pWorksheet)
    {
        assert(pWorksheet);
    }

    virtual ~ExcelWorksheetImpl()
    {
        if (m_pWorksheet)
        {
            m_pWorksheet->Release();
            m_pWorksheet = 0;
        }
    }

    IDispatch* GetIDispatch()
    {
        return m_pWorksheet;
    }

    ELstring   GetName();
    bool       SetName(const ELstring &name);

    ExcelRange GetRange(ELchar columnFrom, ELchar columnTo, int rowFrom, int rowTo);
    ExcelCell  GetCell(ELchar column, int row);

    bool CopyWorksheet(bool after);

private:
    IDispatch *m_pWorksheet;
};


ELstring ExcelWorksheetImpl::GetName()
{
    assert(m_pWorksheet);

    VARIANT result;
    ::VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pWorksheet, DISPATCH_PROPERTYGET, OLESTR("Name"), &result, 0);

    if (FAILED(hr))
        return ELstring();

    ELstring name(result.bstrVal);

    ::VariantClear(&result);

    return name;
}


bool ExcelWorksheetImpl::SetName(const ELstring &name)
{
    assert(m_pWorksheet);

    VARIANT param;
    param.vt = VT_BSTR;
    param.bstrVal = ::SysAllocString(name.c_str());

    HRESULT hr = ComUtil::Invoke(m_pWorksheet, DISPATCH_PROPERTYPUT, OLESTR("Name"), NULL, 1, param);

    ::VariantClear(&param);

    return SUCCEEDED(hr);
}


ExcelRange ExcelWorksheetImpl::GetRange(ELchar columnFrom, ELchar columnTo, int rowFrom, int rowTo)
{
    assert(m_pWorksheet);

    ELchar buf[50];
    memset(buf, 0, sizeof(buf));
    _stprintf_s(buf, 50, ELtext("%c%d:%c%d"), columnFrom, rowFrom, columnTo, rowTo);

    VARIANT param;
    param.vt = VT_BSTR;
    param.bstrVal = ::SysAllocString(buf);

    VARIANT result;
    VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pWorksheet, DISPATCH_PROPERTYGET, OLESTR("Range"), &result, 1, param);

    ::VariantClear(&param);

    if (FAILED(hr))
        return ExcelRange();

    return ExcelRange(result.pdispVal, columnFrom, columnTo, rowFrom, rowTo);
}


ExcelCell ExcelWorksheetImpl::GetCell(ELchar column, int row)
{
    assert(m_pWorksheet);

    ELchar buf[50];
    memset(buf, 0, sizeof(buf));
    _stprintf_s(buf, 50, ELtext("%c%d"), column, row);

    VARIANT param;
    param.vt = VT_BSTR;
    param.bstrVal = ::SysAllocString(buf);

    VARIANT result;
    VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pWorksheet, DISPATCH_PROPERTYGET, OLESTR("Range"), &result, 1, param);

    ::VariantClear(&param);

    if (FAILED(hr))
        return ExcelCell();

    return ExcelCell(result.pdispVal, column, row);
}


bool ExcelWorksheetImpl::CopyWorksheet(bool after)
{
    assert(m_pWorksheet);

    VARIANT paramOptional;
    paramOptional.vt = VT_ERROR;
    paramOptional.scode = DISP_E_PARAMNOTFOUND;

    VARIANT afterParam;
    afterParam.vt = VT_DISPATCH;
    afterParam.pdispVal = m_pWorksheet;

    HRESULT hr;
    
    if (after)
        hr = ComUtil::Invoke(m_pWorksheet, DISPATCH_METHOD, OLESTR("Copy"), NULL, 2, paramOptional, afterParam);
    else
        hr = ComUtil::Invoke(m_pWorksheet, DISPATCH_METHOD, OLESTR("Copy"), NULL, 2, afterParam, paramOptional);

    return SUCCEEDED(hr);
}


////////////////////////////////////////////////////////////////////////////////
// Implementation of class ExcelWorksheet

ExcelWorksheet::ExcelWorksheet(IDispatch *pWorksheet): HandleBase(new ExcelWorksheetImpl(pWorksheet))
{
    assert(pWorksheet);
}


IDispatch* ExcelWorksheet::GetIDispatch()
{
    return Body().GetIDispatch();
}


ELstring ExcelWorksheet::GetName()
{
    return Body().GetName();
}


bool ExcelWorksheet::SetName(const ELstring &name)
{
    return Body().SetName(name);
}


ExcelRange ExcelWorksheet::GetRange(ELchar columnFrom, ELchar columnTo, int rowFrom, int rowTo)
{
    return Body().GetRange(columnFrom, columnTo, rowFrom, rowTo);
}


ExcelCell ExcelWorksheet::GetCell(ELchar column, int row)
{
    return Body().GetCell(column, row);
}


bool ExcelWorksheet::Merge(ELchar columnFrom, ELchar columnTo, int rowFrom, int rowTo, bool multiRow)
{
    ExcelRange range = GetRange(columnFrom, columnTo, rowFrom, rowTo);
    return !range.IsNull() && range.Merge(multiRow);
}


bool ExcelWorksheet::CopyWorksheet(bool after)
{
    return Body().CopyWorksheet(after);
}


// <begin> Handle/Body pattern implementation

ExcelWorksheet::ExcelWorksheet(ExcelWorksheetImpl *impl): HandleBase(impl)
{ 
}


ExcelWorksheetImpl& ExcelWorksheet::Body() const
{
    return dynamic_cast<ExcelWorksheetImpl&>(HandleBase::Body());
}

// <end> Handle/Body pattern implementation


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END
