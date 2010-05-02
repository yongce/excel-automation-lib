/*!
* @file    ExcelWorksheetSet.cpp
* @brief   Implementation file for class ExcelWorksheetSet
* @date    2009-12-13
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#include <tchar.h>
#include <cassert>

#include "ExcelWorksheetSet.h"
#include "ExcelWorksheet.h"
#include "ComUtil.h"
#include "Noncopyable.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


////////////////////////////////////////////////////////////////////////////////
// Definition and implementation of class ExcelWorksheetSetImpl

/*!
* @brief Class ExcelWorksheetSetImpl inplements ExcelWorksheetSet's interfaces.
*/
class ExcelWorksheetSetImpl : public BodyBase, public Noncopyable
{
    // All members are private. Only the friend class ExcelWorksheetSet can access members of ExcelWorksheetSetImpl.
    friend class ExcelWorksheetSet;

private:
    ExcelWorksheetSetImpl(IDispatch *pWorksheetSet): m_pWorksheetSet(pWorksheetSet)
    {
        assert(pWorksheetSet);
    }

    virtual ~ExcelWorksheetSetImpl()
    {
        if (m_pWorksheetSet)
        {
            m_pWorksheetSet->Release();
            m_pWorksheetSet = 0;
        }
    }

    int CountWorksheets();

    ExcelWorksheet GetWorksheet(int index);


private:
    IDispatch *m_pWorksheetSet;
};


int ExcelWorksheetSetImpl::CountWorksheets()
{
    assert(m_pWorksheetSet);

    VARIANT result;
    VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pWorksheetSet, DISPATCH_PROPERTYGET, OLESTR("Count"), &result, 0);

    if (FAILED(hr))
        return -1;

    return (int)result.lVal;
}


ExcelWorksheet ExcelWorksheetSetImpl::GetWorksheet(int index)
{
    assert(m_pWorksheetSet);

    VARIANT param;
    param.vt = VT_INT;
    param.intVal = index;

    VARIANT result;
    VariantInit(&result);

    HRESULT hr = ComUtil::Invoke(m_pWorksheetSet, DISPATCH_PROPERTYGET, OLESTR("Item"), &result, 1, param);

    if (FAILED(hr))
        return ExcelWorksheet();

    return ExcelWorksheet(result.pdispVal);
}


////////////////////////////////////////////////////////////////////////////////
// Implementation of class ExcelWorksheetSet

ExcelWorksheetSet::ExcelWorksheetSet(IDispatch *pWorksheetSet): HandleBase(new ExcelWorksheetSetImpl(pWorksheetSet))
{
    assert(pWorksheetSet);
}


int ExcelWorksheetSet::CountWorksheets()
{
    return Body().CountWorksheets();
}


ExcelWorksheet ExcelWorksheetSet::GetWorksheet(int index)
{
    return Body().GetWorksheet(index);
}


// <begin> Handle/Body pattern implementation

ExcelWorksheetSet::ExcelWorksheetSet(ExcelWorksheetSetImpl *impl): HandleBase(impl)
{ 
}


ExcelWorksheetSetImpl& ExcelWorksheetSet::Body() const
{
    return dynamic_cast<ExcelWorksheetSetImpl&>(HandleBase::Body());
}

// <end> Handle/Body pattern implementation


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END
