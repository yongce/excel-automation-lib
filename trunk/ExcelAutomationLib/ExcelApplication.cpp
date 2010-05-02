/*!
* @file    ExcelApplication.cpp
* @brief   Implementation file for class ExcelApplication
* @date    2009-12-09
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#include <tchar.h>
#include <cassert>

#include "ExcelApplication.h"
#include "ExcelWorkbookSet.h"
#include "ExcelWorkbook.h"
#include "ComUtil.h"
#include "Noncopyable.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


////////////////////////////////////////////////////////////////////////////////
// Definition and implementation of class ExcelWorkbookSetImpl

/*!
* @brief Class ExcelWorkbookSetImpl inplements ExcelWorkbookSet's interfaces.
*/
class ExcelApplicationImpl : public BodyBase, public Noncopyable
{
    // All members are private. Only the friend class ExcelApplication can access members of ExcelApplicationImpl.
    friend class ExcelApplication;

private:
    ExcelApplicationImpl(): m_pApp(0)
    {
        ::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    }

    virtual ~ExcelApplicationImpl()
    {
        if (IsRunning())
            Shutdown();

        CoUninitialize();
    }

    bool IsRunning() const
    {
        return m_pApp != 0;
    }
    
    bool Startup();

    bool SetVisible(bool visible = true);

    bool Shutdown();

    ExcelWorkbookSet GetWorkbookSet();

    ExcelWorkbook OpenWorkbook(const ELchar *filename);

private:
    IDispatch *m_pApp;
    ExcelWorkbookSet m_workbookSet;
};


bool ExcelApplicationImpl::Startup()
{
    assert(!IsRunning());

    CLSID clsid;
	HRESULT hr = ::CLSIDFromProgID(OLESTR("Excel.Application"), &clsid);

	if (FAILED(hr))
		return false;

    IDispatch *pApp = 0;
    hr = ::CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (LPVOID*)&pApp);

    if (FAILED(hr))
        return false;

    m_pApp = pApp;

    return true;
}


bool ExcelApplicationImpl::SetVisible(bool visible /* = true */)
{
    assert(IsRunning());

    VARIANT param;
    param.vt = VT_BOOL;  // 0 == FALSE, -1 == TRUE
    param.boolVal = (visible ? -1 : 0);

    HRESULT hr = ComUtil::Invoke(m_pApp, DISPATCH_PROPERTYPUT, OLESTR("Visible"), NULL, 1, param);

    return SUCCEEDED(hr);
}


bool ExcelApplicationImpl::Shutdown()
{
    assert(IsRunning());

    HRESULT hr = ComUtil::Invoke(m_pApp, DISPATCH_METHOD, OLESTR("Quit"), NULL, 0);
    m_pApp->Release();
    m_pApp = 0;
    return SUCCEEDED(hr);
}


ExcelWorkbookSet ExcelApplicationImpl::GetWorkbookSet()
{
    assert(IsRunning());

    if (!m_workbookSet.IsNull())
        return m_workbookSet;

    VARIANT result;
    VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pApp, DISPATCH_PROPERTYGET, OLESTR("Workbooks"), &result, 0);

    if (SUCCEEDED(hr))
        m_workbookSet = ExcelWorkbookSet(result.pdispVal);

    return m_workbookSet;
}


ExcelWorkbook ExcelApplicationImpl::OpenWorkbook(const ELchar *filename)
{
    assert(IsRunning());

    if (m_workbookSet.IsNull())
        GetWorkbookSet();

    return m_workbookSet.OpenWorkbook(filename);
}


////////////////////////////////////////////////////////////////////////////////
// Implementation of class ExcelWorkbookSet

ExcelApplication::ExcelApplication(): HandleBase(new ExcelApplicationImpl())
{
}


bool ExcelApplication::IsRunning() const
{
    return Body().IsRunning();
}


bool ExcelApplication::Startup()
{
    return Body().Startup();
}


bool ExcelApplication::SetVisible(bool visible /* = true */)
{
    return Body().SetVisible(visible);
}


bool ExcelApplication::Shutdown()
{
    return Body().Shutdown();
}


ExcelWorkbook ExcelApplication::OpenWorkbook(const ELchar *filename)
{
    return Body().OpenWorkbook(filename);
}


// <begin> Handle/Body pattern implementation

ExcelApplication::ExcelApplication(ExcelApplicationImpl *impl): HandleBase(impl)
{ 
}


ExcelApplicationImpl& ExcelApplication::Body() const
{
    return dynamic_cast<ExcelApplicationImpl&>(HandleBase::Body());
}

// <end> Handle/Body pattern implementation


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END
