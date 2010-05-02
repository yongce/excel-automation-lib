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

    ELstring GetName();

    ExcelRange GetRange(ELchar columnFrom, ELchar columnTo, int rowFrom, int rowTo);
    ExcelCell  GetCell(ELchar column, int row);


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


////////////////////////////////////////////////////////////////////////////////
// Implementation of class ExcelWorksheet

ExcelWorksheet::ExcelWorksheet(IDispatch *pWorksheet): HandleBase(new ExcelWorksheetImpl(pWorksheet))
{
    assert(pWorksheet);
}


ELstring ExcelWorksheet::GetName()
{
    return Body().GetName();
}


ExcelRange ExcelWorksheet::GetRange(ELchar columnFrom, ELchar columnTo, int rowFrom, int rowTo)
{
    return Body().GetRange(columnFrom, columnTo, rowFrom, rowTo);
}


ExcelCell ExcelWorksheet::GetCell(ELchar column, int row)
{
    return Body().GetCell(column, row);
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
