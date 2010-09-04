/*!
* @file    ExcelFont.cpp
* @brief   Implementation file for class ExcelFont
* @date    2010-09-01 15:12:42
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#include "ExcelFont.h"
#include "ComUtil.h"
#include "Noncopyable.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


////////////////////////////////////////////////////////////////////////////////
// Definition and implementation of class ExcelFontImpl

/*!
* @brief Class ExcelFontImpl inplements ExcelFont's interfaces.
*/
class ExcelFontImpl : public BodyBase, public Noncopyable
{
    // All members are private. Only the friend class ExcelFont can access members of ExcelFontImpl.
    friend class ExcelFont;

private:
    ExcelFontImpl(IDispatch *pFont) : m_pFont(pFont)
    {
        assert(pFont);
    }

    virtual ~ExcelFontImpl()
    {
        if (m_pFont)
        {
            m_pFont->Release();
            m_pFont = 0;
        }
    }

    bool GetName(ELstring &name);
    bool SetName(const ELstring &name);

    bool GetSize(int &size);
    bool SetSize(int size);

    bool GetBold(bool &bold);
    bool SetBold(bool bold);

    bool GetItalic(bool &italic);
    bool SetItalic(bool italic);

    bool GetColor(COLORREF &rgb);
    bool SetColor(COLORREF rgb);


private:
    IDispatch *m_pFont;
};


bool ExcelFontImpl::GetName(ELstring &name)
{
    assert(m_pFont);

    VARIANT result;
    ::VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pFont, DISPATCH_PROPERTYGET, OLESTR("Name"), &result, 0);

    if (SUCCEEDED(hr))
        name = result.bstrVal;

    ::VariantClear(&result);

    return SUCCEEDED(hr);
}


bool ExcelFontImpl::SetName(const ExcelAutomation::ELstring &name)
{
    assert(m_pFont);

    VARIANT param;
    param.vt = VT_BSTR;
    param.bstrVal = ::SysAllocString(name.c_str());

    HRESULT hr = ComUtil::Invoke(m_pFont, DISPATCH_PROPERTYPUT, OLESTR("Name"), NULL, 1, param);

    ::VariantClear(&param);

    return SUCCEEDED(hr);
}


bool ExcelFontImpl::GetSize(int &size)
{
    assert(m_pFont);

    VARIANT result;
    ::VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pFont, DISPATCH_PROPERTYGET, OLESTR("Size"), &result, 0);

    if (SUCCEEDED(hr))
    {
        hr = ::VariantChangeType(&result, &result, VARIANT_NOUSEROVERRIDE, VT_INT);
        if (SUCCEEDED(hr))
            size = result.intVal;
    }

    return SUCCEEDED(hr);
}


bool ExcelFontImpl::SetSize(int size)
{
    assert(m_pFont);

    VARIANT param;
    param.vt = VT_INT;
    param.intVal = size;

    HRESULT hr = ComUtil::Invoke(m_pFont, DISPATCH_PROPERTYPUT, OLESTR("Size"), NULL, 1, param);

    return SUCCEEDED(hr); 
}


bool ExcelFontImpl::GetBold(bool &bold)
{
    assert(m_pFont);

    VARIANT result;
    ::VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pFont, DISPATCH_PROPERTYGET, OLESTR("Bold"), &result, 0);

    if (SUCCEEDED(hr))
    {
        hr = ::VariantChangeType(&result, &result, VARIANT_NOUSEROVERRIDE, VT_BOOL);
        if (SUCCEEDED(hr))
            bold = (result.boolVal == TRUE);
    }

    return SUCCEEDED(hr);
}


bool ExcelFontImpl::SetBold(bool bold)
{
    assert(m_pFont);

    VARIANT param;
    param.vt = VT_BOOL;
    param.boolVal = bold;

    HRESULT hr = ComUtil::Invoke(m_pFont, DISPATCH_PROPERTYPUT, OLESTR("Bold"), NULL, 1, param);

    return SUCCEEDED(hr); 
}


bool ExcelFontImpl::GetItalic(bool &italic)
{
    assert(m_pFont);

    VARIANT result;
    ::VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pFont, DISPATCH_PROPERTYGET, OLESTR("Italic"), &result, 0);

    if (SUCCEEDED(hr))
    {
        hr = ::VariantChangeType(&result, &result, VARIANT_NOUSEROVERRIDE, VT_BOOL);
        if (SUCCEEDED(hr))
            italic = (result.boolVal == TRUE);
    }

    return SUCCEEDED(hr);
}


bool ExcelFontImpl::SetItalic(bool italic)
{
    assert(m_pFont);

    VARIANT param;
    param.vt = VT_BOOL;
    param.boolVal = italic;

    HRESULT hr = ComUtil::Invoke(m_pFont, DISPATCH_PROPERTYPUT, OLESTR("Italic"), NULL, 1, param);

    return SUCCEEDED(hr); 
}


bool ExcelFontImpl::GetColor(COLORREF &rgb)
{
    assert(m_pFont);

    VARIANT result;
    ::VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pFont, DISPATCH_PROPERTYGET, OLESTR("Color"), &result, 0);

    if (SUCCEEDED(hr))
    {
        hr = ::VariantChangeType(&result, &result, VARIANT_NOUSEROVERRIDE, VT_UI4);
        if (SUCCEEDED(hr))
            rgb = result.ulVal;
    }

    return SUCCEEDED(hr);
}


bool ExcelFontImpl::SetColor(COLORREF rgb)
{
    assert(m_pFont);

    VARIANT param;
    param.vt = VT_UI4;
    param.ulVal = rgb;

    HRESULT hr = ComUtil::Invoke(m_pFont, DISPATCH_PROPERTYPUT, OLESTR("Color"), NULL, 1, param);

    return SUCCEEDED(hr); 
}


////////////////////////////////////////////////////////////////////////////////
// Implementation of class ExcelFont


ExcelFont::ExcelFont(IDispatch *pFont) : HandleBase(new ExcelFontImpl(pFont))
{
    assert(pFont);
}


bool ExcelFont::GetName(ELstring &name)
{
    return Body().GetName(name);
}


bool ExcelFont::SetName(const ELstring &name)
{
    return Body().SetName(name);
}


bool ExcelFont::GetSize(int &size)
{
    return Body().GetSize(size);
}


bool ExcelFont::SetSize(int size)
{
    return Body().SetSize(size);
}


bool ExcelFont::GetBold(bool &bold)
{
    return Body().GetBold(bold);
}


bool ExcelFont::SetBold(bool bold)
{
    return Body().SetBold(bold);
}


bool ExcelFont::GetItalic(bool &italic)
{
    return Body().GetItalic(italic);
}


bool ExcelFont::SetItalic(bool italic)
{
    return Body().SetItalic(italic);
}


bool ExcelFont::GetColor(COLORREF &rgb)
{
    return Body().GetColor(rgb);
}


bool ExcelFont::SetColor(COLORREF rgb)
{
    return Body().SetColor(rgb);
}


// <begin> Handle/Body pattern implementation

ExcelFont::ExcelFont(ExcelFontImpl *impl): HandleBase(impl)
{ 
}


ExcelFontImpl& ExcelFont::Body() const
{
    return dynamic_cast<ExcelFontImpl&>(HandleBase::Body());
}

// <end> Handle/Body pattern implementation


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END

