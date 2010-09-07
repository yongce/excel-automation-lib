/*!
* @file    ExcelWorkbookSet.h
* @brief   Header file for class ExcelWorkbookSet
* @date    2009-12-01
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef EXCELWORKBOOKSET_H_GUID_728B1E14_CEA1_4602_983A_652E315FF824
#define EXCELWORKBOOKSET_H_GUID_728B1E14_CEA1_4602_983A_652E315FF824


#include "LibDef.h"
#include "HandleBody.h"
#include "StringUtil.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


// Forward declaration
class ExcelWorkbook;


/*!
* @brief Class ExcelWorkbookSet represents the concept "Workbooks" in Excel.
* @note ExcelWorkbookSet/ExcelWorkbookSetImpl is an implementation of the "Handle/Body" pattern.
*/
class EXCEL_AUTOMATION_DLL_API ExcelWorkbookSet : public HandleBase
{
public:
    /*!
    * Default constructor
    */ // Doc is needed by Doxygen
    ExcelWorkbookSet(): HandleBase(0) { }

    ExcelWorkbook OpenWorkbook(const ELchar *filename);

    ExcelWorkbook CreateWorkbook(const ELchar *filename);

private:
    friend class ExcelApplicationImpl;    // which will call the following ctor
    ExcelWorkbookSet(IDispatch *pWorkbookSet);

private:
    // <begin> Handle/Body pattern implementation
    friend class ExcelWorkbookSetImpl;
    ExcelWorkbookSet(ExcelWorkbookSetImpl *impl);
    ExcelWorkbookSetImpl& Body() const;
    // <end> Handle/Body pattern implementation
};


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END


#endif //EXCELWORKBOOKSET_H_GUID_728B1E14_CEA1_4602_983A_652E315FF824
