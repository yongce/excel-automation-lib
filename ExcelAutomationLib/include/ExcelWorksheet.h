/*!
* @file    ExcelWorksheet.h
* @brief   Header file for class ExcelWorksheet
* @date    2009-12-01
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef EXCELWORKSHEET_H_GUID_61B8B170_8EC6_4530_8CB9_E4B017D81BC0
#define EXCELWORKSHEET_H_GUID_61B8B170_8EC6_4530_8CB9_E4B017D81BC0


#include "LibDef.h"
#include "HandleBody.h"
#include "StringUtil.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


// Forward declaration
class ExcelRange;
class ExcelCell;


/*!
* @brief Class ExcelWorksheet represents the concept "Worksheet" in Excel.
* @note ExcelWorksheet/ExcelWorksheetImpl is an implementation of the "Handle/Body" pattern.
*/
class EXCEL_AUTOMATION_DLL_API ExcelWorksheet : public HandleBase
{
public:
    /*!
    * Default constructor
    */ // Doc is needed by Doxygen
    ExcelWorksheet(): HandleBase(0) { }

    ELstring   GetName();
    ExcelRange GetRange(ELchar columnFrom, ELchar columnTo, int rowFrom, int rowTo);
    ExcelCell  GetCell(ELchar column, int row);

private:
    friend class ExcelWorkbookImpl;      // which calls the following ctor
    friend class ExcelWorksheetSetImpl;  // which calls the following ctor
    ExcelWorksheet(IDispatch *pWorksheet);

private:
    // <begin> Handle/Body pattern implementation
    friend class ExcelWorksheetImpl;
    ExcelWorksheet(ExcelWorksheetImpl *impl);
    ExcelWorksheetImpl& Body() const;
    // <end> Handle/Body pattern implementation
};


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END


#endif //EXCELWORKSHEET_H_GUID_61B8B170_8EC6_4530_8CB9_E4B017D81BC0
