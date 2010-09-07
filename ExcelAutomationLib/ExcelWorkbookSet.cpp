/*!
* @file    ExcelWorkbookSet.cpp
* @brief   Implementation file for class ExcelWorkbookSet
* @date    2009-12-01
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#include <tchar.h>
#include <cassert>
#include <vector>

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
class ExcelWorkbookSetImpl : public BodyBase, public Noncopyable
{
    // All members are private. Only the friend class ExcelWorkbookSet can access members of ExcelWorkbookSetImpl.
    friend class ExcelWorkbookSet;

private:
    ExcelWorkbookSetImpl(IDispatch *pWorkbookSet): m_pWorkbookSet(pWorkbookSet)
    {
        assert(pWorkbookSet);
    }

    virtual ~ExcelWorkbookSetImpl()
    {
        if (m_pWorkbookSet)
        {
            m_pWorkbookSet->Release();
            m_pWorkbookSet = 0;
        }
    }

    ExcelWorkbook OpenWorkbook(const ELchar *filename);

    ExcelWorkbook CreateWorkbook(const ELchar *filename);


private:
    IDispatch *m_pWorkbookSet;
};


ExcelWorkbook ExcelWorkbookSetImpl::OpenWorkbook(const ELchar *filename)
{
    assert(m_pWorkbookSet);

    // Get full path name for the file
    std::vector<ELchar> fullpath(256, ELtext('\0'));
    DWORD bufLen = ::GetFullPathName(filename, fullpath.size(), &fullpath[0], 0);
    if (bufLen == 0)
        return ExcelWorkbook();  // the file cannot be found

    if (bufLen > fullpath.size())
    {
        fullpath.resize(bufLen, ELtext('\0'));
        ::GetFullPathName(filename, fullpath.size(), &fullpath[0], 0);
    }

    VARIANT param;
    param.vt = VT_BSTR;
    param.bstrVal = ::SysAllocString(&fullpath[0]);

    VARIANT result;
    VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pWorkbookSet, DISPATCH_METHOD, OLESTR("Open"), &result, 1, param);

    VariantClear(&param);

    if (FAILED(hr))
        return ExcelWorkbook();

    return ExcelWorkbook(result.pdispVal);
}


ExcelWorkbook ExcelWorkbookSetImpl::CreateWorkbook(const ELchar *filename)
{
    assert(m_pWorkbookSet);

    VARIANT result;
    VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pWorkbookSet, DISPATCH_METHOD, OLESTR("Add"), &result, 0);

    if (FAILED(hr))
        return ExcelWorkbook();

    ExcelWorkbook workbook(result.pdispVal);
    if (!workbook.SaveAs(filename))
        return ExcelWorkbook();

    return workbook;
}


////////////////////////////////////////////////////////////////////////////////
// Implementation of class ExcelWorkbookSet

ExcelWorkbookSet::ExcelWorkbookSet(IDispatch *pWorkbookSet): HandleBase(new ExcelWorkbookSetImpl(pWorkbookSet))
{
    assert(pWorkbookSet);
}


ExcelWorkbook ExcelWorkbookSet::OpenWorkbook(const ELchar *filename)
{
    return Body().OpenWorkbook(filename);
}


ExcelWorkbook ExcelWorkbookSet::CreateWorkbook(const ELchar *filename)
{
    return Body().CreateWorkbook(filename);
}


// <begin> Handle/Body pattern implementation

ExcelWorkbookSet::ExcelWorkbookSet(ExcelWorkbookSetImpl *impl): HandleBase(impl)
{ 
}


ExcelWorkbookSetImpl& ExcelWorkbookSet::Body() const
{
    return dynamic_cast<ExcelWorkbookSetImpl&>(HandleBase::Body());
}

// <end> Handle/Body pattern implementation


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END
