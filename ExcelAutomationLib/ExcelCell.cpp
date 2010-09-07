/*!
* @file    ExcelCell.cpp
* @brief   Implementation file for class ExcelCell
* @date    2009-12-31
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#include "ExcelCell.h"
#include "ComUtil.h"
#include "Noncopyable.h"
#include "ExcelFont.h"
#include "ExcelUtil.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


////////////////////////////////////////////////////////////////////////////////
// Definition and implementation of class ExcelCellImpl

/*!
* @brief Class ExcelCellImpl inplements ExcelCell's interfaces.
*/
class ExcelCellImpl : public BodyBase, public Noncopyable
{
    // All members are private. Only the friend class ExcelCell can access members of ExcelCellImpl.
    friend class ExcelCell;

private:
    ExcelCellImpl(IDispatch *pCell, ELchar column, int row): m_pCell(pCell), m_column(column), m_row(row)
    {
        assert(pCell);
    }

    virtual ~ExcelCellImpl()
    {
        if (m_pCell)
        {
            m_pCell->Release();
            m_pCell = 0;
        }
    }

    bool GetValue(ELstring &value);
    bool SetValue(const ELstring &value);
    bool SetValue(int value);
    bool SetValue(double value);

    ExcelFont GetFont();

    bool SetHorizontalAlignment(ExcelHorizontalAlignment align);
    bool SetVerticalAlignment(ExcelVerticalAlignment align);

private:
    IDispatch *m_pCell;      // in fact, it refers an "Range" object
    ELchar     m_column;
    int        m_row;
};


bool ExcelCellImpl::GetValue(ELstring &value)
{
    assert(m_pCell);

    VARIANT result;
    ::VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pCell, DISPATCH_PROPERTYGET, OLESTR("Value"), &result, 0);

    if (SUCCEEDED(hr))
    {
        hr = ::VariantChangeType(&result, &result, VARIANT_NOUSEROVERRIDE, VT_BSTR);
        if (SUCCEEDED(hr))
            value = result.bstrVal;
    }

    ::VariantClear(&result);

    return SUCCEEDED(hr);
}


bool ExcelCellImpl::SetValue(const ELstring &value)
{
    assert(m_pCell);

    VARIANT param;
    param.vt = VT_BSTR;
    param.bstrVal = ::SysAllocString(value.c_str());

    HRESULT hr = ComUtil::Invoke(m_pCell, DISPATCH_PROPERTYPUT, OLESTR("Value"), NULL, 1, param);

    ::VariantClear(&param);

    return SUCCEEDED(hr);
}


bool ExcelCellImpl::SetValue(int value)
{
    assert(m_pCell);

    VARIANT param;
    param.vt = VT_INT;
    param.intVal = value;

    HRESULT hr = ComUtil::Invoke(m_pCell, DISPATCH_PROPERTYPUT, OLESTR("Value"), NULL, 1, param);

    return SUCCEEDED(hr);
}


bool ExcelCellImpl::SetValue(double value)
{
    assert(m_pCell);

    VARIANT param;
    param.vt = VT_R8;
    param.dblVal = value;

    HRESULT hr = ComUtil::Invoke(m_pCell, DISPATCH_PROPERTYPUT, OLESTR("Value"), NULL, 1, param);

    return SUCCEEDED(hr);
}


ExcelFont ExcelCellImpl::GetFont()
{
    assert(m_pCell);

    VARIANT result;
    ::VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pCell, DISPATCH_PROPERTYGET, OLESTR("Font"), &result, 0);

    if (FAILED(hr))
        return ExcelFont();

    return ExcelFont(result.pdispVal);
}


bool ExcelCellImpl::SetHorizontalAlignment(ExcelHorizontalAlignment align)
{
    assert(m_pCell);

    int alignConstant;
    if (!ExcelUtil::GetExcelConstant(align, alignConstant))
        return false;

    VARIANT param;
    param.vt = VT_INT;
    param.intVal = alignConstant;

    HRESULT hr = ComUtil::Invoke(m_pCell, DISPATCH_PROPERTYPUT, OLESTR("HorizontalAlignment"), NULL, 1, param);

    return SUCCEEDED(hr);
}


bool ExcelCellImpl::SetVerticalAlignment(ExcelVerticalAlignment align)
{
    assert(m_pCell);

    int alignConstant;
    if (!ExcelUtil::GetExcelConstant(align, alignConstant))
        return false;

    VARIANT param;
    param.vt = VT_INT;
    param.intVal = alignConstant;

    HRESULT hr = ComUtil::Invoke(m_pCell, DISPATCH_PROPERTYPUT, OLESTR("VerticalAlignment"), NULL, 1, param);

    return SUCCEEDED(hr);    
}


////////////////////////////////////////////////////////////////////////////////
// Implementation of class ExcelCell

ExcelCell::ExcelCell(IDispatch *pCell, ELchar column, int row): HandleBase(new ExcelCellImpl(pCell, column, row))
{
    assert(pCell);
}


bool ExcelCell::GetValue(ELstring &value)
{
    return Body().GetValue(value);
}


bool ExcelCell::SetValue(const ELstring &value)
{
    return Body().SetValue(value);
}


bool ExcelCell::SetValue(int value)
{
    return Body().SetValue(value);
}


bool ExcelCell::SetValue(double value)
{
    return Body().SetValue(value);
}


ExcelFont ExcelCell::GetFont()
{
    return Body().GetFont();
}


bool ExcelCell::SetHorizontalAlignment(ExcelHorizontalAlignment align)
{
    return Body().SetHorizontalAlignment(align);
}


bool ExcelCell::SetVerticalAlignment(ExcelVerticalAlignment align)
{
    return Body().SetVerticalAlignment(align);
}


// <begin> Handle/Body pattern implementation

ExcelCell::ExcelCell(ExcelCellImpl *impl): HandleBase(impl)
{ 
}


ExcelCellImpl& ExcelCell::Body() const
{
    return dynamic_cast<ExcelCellImpl&>(HandleBase::Body());
}

// <end> Handle/Body pattern implementation


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END

