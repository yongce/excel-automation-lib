/*!
* @file    ExcelWorksheetSet.h
* @brief   Header file for class ExcelWorksheetSet
* @date    2009-12-13
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef EXCELWORKSHEETSET_H_GUID_5BE4FB5C_0B00_47B7_916C_6505FF383D20
#define EXCELWORKSHEETSET_H_GUID_5BE4FB5C_0B00_47B7_916C_6505FF383D20


#include "LibDef.h"
#include "HandleBody.h"
#include "StringUtil.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


// Forward declaration
class ExcelWorksheet;


/*!
* @brief Class ExcelWorksheetSet represents the concept "Worksheets" in Excel.
* @note ExcelWorksheetSet/ExcelWorksheetSetImpl is an implementation of the "Handle/Body" pattern.
*/
class EXCEL_AUTOMATION_DLL_API ExcelWorksheetSet : public HandleBase
{
public:
    /*!
    * Default constructor
    */ // Doc is needed by Doxygen
    ExcelWorksheetSet(): HandleBase(0) { }

    int CountWorksheets();

    /*!
    * @param [in] index Starts from 1
    */
    ExcelWorksheet GetWorksheet(int index);

private:
    friend class ExcelWorkbookImpl;   // which calls the following ctor
    ExcelWorksheetSet(IDispatch *pWorksheetSet);

private:
    // <begin> Handle/Body pattern implementation
    friend class ExcelWorksheetSetImpl;
    ExcelWorksheetSet(ExcelWorksheetSetImpl *impl);
    ExcelWorksheetSetImpl& Body() const;
    // <end> Handle/Body pattern implementation
};


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END


#endif //EXCELWORKSHEETSET_H_GUID_5BE4FB5C_0B00_47B7_916C_6505FF383D20
