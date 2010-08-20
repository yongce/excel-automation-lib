/*!
* @file    ExcelWorkbook.h
* @brief   Header file for class ExcelWorkbook
* @date    2009-12-01
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef EXCELWORKBOOK_H_GUID_FFB67A8F_9104_4F1F_B552_B9EF17C5DD4E
#define EXCELWORKBOOK_H_GUID_FFB67A8F_9104_4F1F_B552_B9EF17C5DD4E


#include "LibDef.h"
#include "HandleBody.h"
#include "StringUtil.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


// Forward declaration
class ExcelWorksheet;
class ExcelWorksheetSet;


/*!
* @brief Class ExcelWorkbook represents the concept "Workbook" in Excel.
* @note ExcelWorkbook/ExcelWorkbookImpl is an implementation of the "Handle/Body" pattern.
*/
class EXCEL_AUTOMATION_DLL_API ExcelWorkbook : public HandleBase
{
public:
    /*!
    * Default constructor
    */ // Doc is needed by Doxygen
    ExcelWorkbook(): HandleBase(0) { }

    ExcelWorksheet GetActiveWorksheet() const;
    ExcelWorksheetSet GetAllWorksheets() const;

    bool Save() const;
    bool SaveAs(const ELstring &filename);

    bool Close() const;

private:
    friend class ExcelWorkbookSetImpl;   // which will call the following ctor
    ExcelWorkbook(IDispatch *pWorkbook);

private:
    // <begin> Handle/Body pattern implementation
    friend class ExcelWorkbookImpl;
    ExcelWorkbook(ExcelWorkbookImpl *impl);
    ExcelWorkbookImpl& Body() const;
    // <end> Handle/Body pattern implementation
};


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END


#endif //EXCELWORKBOOK_H_GUID_FFB67A8F_9104_4F1F_B552_B9EF17C5DD4E
