/*!
* @file    ExcelRange.cpp
* @brief   Implementation file for class ExcelRange
* @date    2009-12-08
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#include <tchar.h>
#include <cassert>
#include <sstream>

#include "ExcelRange.h"
#include "StringUtil.h"
#include "ComUtil.h"
#include "Noncopyable.h"
#include "ExcelFont.h"
#include "ExcelUtil.h"


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
        m_pRange(pRange), m_columnFrom(columnFrom), m_columnTo(columnTo), m_rowFrom(rowFrom), m_rowTo(rowTo),
        m_merged(false), m_multiRowMerged(false)
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

    bool Merge(bool multiRow);

    ExcelFont GetFont();

    bool SetHorizontalAlignment(ExcelHorizontalAlignment align);
    bool SetVerticalAlignment(ExcelVerticalAlignment align);
    

private:
    IDispatch *m_pRange;
    ELchar     m_columnFrom;
    ELchar     m_columnTo;
    int        m_rowFrom;
    int        m_rowTo;
    bool       m_merged;
    bool       m_multiRowMerged;
};


bool ExcelRangeImpl::ReadData(ELstring &data)
{
    assert(!m_merged);
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
    assert(!m_merged);
    assert(m_pRange);

    VARIANT param;
    param.vt = VT_ARRAY | VT_VARIANT;
    param.parray = ComUtil::DecodeSafeArrayDim2(data);

    HRESULT hr = ComUtil::Invoke(m_pRange, DISPATCH_PROPERTYPUT, OLESTR("Value"), NULL, 1, param);

    ::VariantClear(&param);

    return SUCCEEDED(hr);
}


bool ExcelRangeImpl::Merge(bool multiRow)
{
    assert(m_pRange);

    VARIANT param;
    param.vt = VT_BOOL;
    param.boolVal = multiRow;

    HRESULT hr = ComUtil::Invoke(m_pRange, DISPATCH_METHOD, OLESTR("Merge"), NULL, 1, param);

    m_merged = SUCCEEDED(hr);
    m_multiRowMerged = multiRow;

    return SUCCEEDED(hr);
}


ExcelFont ExcelRangeImpl::GetFont()
{
    assert(m_pRange);

    VARIANT result;
    ::VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pRange, DISPATCH_PROPERTYGET, OLESTR("Font"), &result, 0);

    if (FAILED(hr))
        return ExcelFont();

    return ExcelFont(result.pdispVal);
}


bool ExcelRangeImpl::SetHorizontalAlignment(ExcelHorizontalAlignment align)
{
    assert(m_pRange);

    int alignConstant;
    if (!ExcelUtil::GetExcelConstant(align, alignConstant))
        return false;

    VARIANT param;
    param.vt = VT_INT;
    param.intVal = alignConstant;

    HRESULT hr = ComUtil::Invoke(m_pRange, DISPATCH_PROPERTYPUT, OLESTR("HorizontalAlignment"), NULL, 1, param);

    return SUCCEEDED(hr);
}


bool ExcelRangeImpl::SetVerticalAlignment(ExcelVerticalAlignment align)
{
    assert(m_pRange);

    int alignConstant;
    if (!ExcelUtil::GetExcelConstant(align, alignConstant))
        return false;

    VARIANT param;
    param.vt = VT_INT;
    param.intVal = alignConstant;

    HRESULT hr = ComUtil::Invoke(m_pRange, DISPATCH_PROPERTYPUT, OLESTR("VerticalAlignment"), NULL, 1, param);

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


bool ExcelRange::ReadData(std::vector<std::vector<ELstring> > &values)
{
    ELstring tmp;
    if (!ReadData(tmp))
        return false;

    return DecodeData(tmp, values);
}


bool ExcelRange::WriteData(const ELchar *data)
{
    return Body().WriteData(data);
}


bool ExcelRange::WriteData(const ELstring &data)
{
    return WriteData(data.c_str());
}


bool ExcelRange::WriteData(const std::vector<std::vector<ELstring> > &values)
{
    ELstring tmp = EncodeData(values);
    return WriteData(tmp);
}


bool ExcelRange::DecodeData(const ELstring &data, std::vector<std::vector<ELstring> > &values)
{
    std::basic_istringstream<ELchar> iss(data);
    iss >> std::noskipws;

    int row = 0;
    int column = 0;
    ELchar dumb;

    // Encoding format: <row>#<column>#
    iss >> row >> dumb;             // <row>#
    assert(dumb == ELtext('#'));    // Whether the delimiter is '#', it's not important. 
                                    // So we just place an assert statement here.
    iss >> column >> dumb;          // <column>#
    assert(dumb == ELtext('#'));

    if (row <=0 || column <= 0)
        return false;   // no data or dirty data

    // initialize the two-dimensional array
    std::vector<std::vector<ELstring> >(row, std::vector<ELstring>(column, ELstring())).swap(values);

    bool validState = true;  // flag indicating whether the data is well formed

    for (int i = 0; validState && i < row; ++i)
    {
        for (int j = 0; validState && j < column; ++j)
        {
            // Encoding format: <number of characters>#<characters>
            int count = 0;
            iss >> count >> dumb;
            assert(dumb == ELtext('#'));

            validState = iss.good() && (count >= 0);

            ELstring curValue;
            for (int k = 0; k < count; ++k)
            {
                ELchar ch;
                iss >> ch;
                curValue.push_back(ch);
            }

            validState = validState && iss.good();

            if (validState)
                values[i][j] = curValue;
        }
    }

    return validState;
}


ELstring ExcelRange::EncodeData(const std::vector<std::vector<ELstring> > &values)
{
    std::basic_ostringstream<ELchar> oss;

    int rowNum = values.size();
    if (rowNum == 0)
        return ELstring(ELtext("0#"));

    int columnNum = values[0].size();

    // Encoding format: <row>#<column>#
    oss << rowNum << ELtext('#') << columnNum << ELtext('#');

    for (int i = 0; i < rowNum; ++i)
    {
        for (int j = 0; j < columnNum; ++j)
        {
            // Encoding format: <number of characters>#<characters>
            oss << values[i][j].length() << ELtext('#') << values[i][j];
        }
    }

    return oss.str();
}


bool ExcelRange::Merge(bool multiRow /* = false */)
{
    return Body().Merge(multiRow);
}


ExcelFont ExcelRange::GetFont()
{
    return Body().GetFont();
}


bool ExcelRange::SetHorizontalAlignment(ExcelHorizontalAlignment align)
{
    return Body().SetHorizontalAlignment(align);
}


bool ExcelRange::SetVerticalAlignment(ExcelVerticalAlignment align)
{
    return Body().SetVerticalAlignment(align);
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
