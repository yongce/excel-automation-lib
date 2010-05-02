/*!
* @file    ExcelRange.cpp
* @brief   Implementation file for class ExcelRange
* @date    2009-12-08
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#include <tchar.h>
#include <cassert>

#include "ExcelRange.h"
#include "StringUtil.h"
#include "ComUtil.h"
#include "Noncopyable.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START

////////////////////////////////////////////////////////////////////////////////
// Definition and implementation of class ExcelRangeImpl

/*!
* @brief Class ExcelWorksheetImpl inplements ExcelWorksheet's interfaces.
* @note: All calling for ExcelWorksheet's interface will redirect to ExcelRangeImpl.
*/
class ExcelRangeImpl : public BodyBase, public Noncopyable
{
    // All members are private, so only the friend class ExcelRange can access the members of ExcelRangeImpl
    friend class ExcelRange;

private:
    ExcelRangeImpl(IDispatch *pRange, ELchar columnFrom, ELchar columnTo, int rowFrom, int rowTo): \
        m_pRange(pRange), m_columnFrom(columnFrom), m_columnTo(columnTo), m_rowFrom(rowFrom), m_rowTo(rowTo)
    {
        assert(pRange);
    }

    virtual ~ExcelRangeImpl()
    {
        if (m_pRange)
        {
            m_pRange->Release();
            m_pRange = 0;
        }
    }

    bool ReadData(ELstring &data);
    bool WriteData(const ELchar *data);
    

private:
    IDispatch *m_pRange;
    ELchar     m_columnFrom;
    ELchar     m_columnTo;
    int        m_rowFrom;
    int        m_rowTo;
};


bool ExcelRangeImpl::ReadData(ELstring &data)
{
    assert(m_pRange);
    data.clear();

    VARIANT result;
    ::VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pRange, DISPATCH_PROPERTYGET, OLESTR("Value"), &result, 0);

    if (SUCCEEDED(hr))
    {
        hr = ComUtil::EncodeSafeArrayDim2(result.parray, data);
        ::VariantClear(&result);
    }

    return SUCCEEDED(hr);
}



bool ExcelRangeImpl::WriteData(const ELchar *data)
{
    assert(m_pRange);

    VARIANT param;
    param.vt = VT_ARRAY | VT_VARIANT;
    param.parray = ComUtil::DecodeSafeArrayDim2(data);

    HRESULT hr = ComUtil::Invoke(m_pRange, DISPATCH_PROPERTYPUT, OLESTR("Value"), NULL, 1, param);

    ::VariantClear(&param);

    return SUCCEEDED(hr);
}


////////////////////////////////////////////////////////////////////////////////
// class ExcelRange implementation

ExcelRange::ExcelRange(IDispatch *pRange, ELchar columnFrom, ELchar columnTo, int rowFrom, int rowTo): 
    HandleBase(new ExcelRangeImpl(pRange, columnFrom, columnTo, rowFrom, rowTo))
{
    assert(pRange);
}


bool ExcelRange::ReadData(ELstring &data)
{
    return Body().ReadData(data);
}


bool ExcelRange::WriteData(const ELchar *data)
{
    return Body().WriteData(data);
}


// <begin> Handle/Body pattern implementation

ExcelRange::ExcelRange(ExcelRangeImpl *impl): HandleBase(impl)
{
}


ExcelRangeImpl& ExcelRange::Body() const
{
    return dynamic_cast<ExcelRangeImpl&>(HandleBase::Body());
}

// <end> Handle/Body pattern implementation



// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END
