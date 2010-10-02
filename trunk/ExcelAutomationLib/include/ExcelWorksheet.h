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

    /*!
    * @brief Get the IDispatch pointer for this worksheet object
    * @note This pointer will go to be invalid when this ExcelWorksheet object is destroyed.
    * @note Don't call IDispatch::Release() on this IDispatch pointer.
    */
    IDispatch* GetIDispatch(); 

    ELstring   GetName();
    bool       SetName(const ELstring &name);

    ExcelRange GetRange(ELchar columnFrom, ELchar columnTo, int rowFrom, int rowTo);
    ExcelCell  GetCell(ELchar column, int row);

    /*!
    * @brief Merge the specified range into one cell or merge every row of the range into one cell.
    * @param [in] columnFrom Left column of the range
    * @param [in] columnTo Right column of the range
    * @param [in] rowFrom Top row of the range
    * @param [in] rowTo Bottom row of the range
    * @param [in] multiRow If true, merge every row of the range into one cell;
    *                      otherwise (false), merge the whole range into one cell.
    * @return true if successful, otherwise false
    */
    bool Merge(ELchar columnFrom, ELchar columnTo, int rowFrom, int rowTo, bool multiRow = false);

    /*!
    * @brief Create a copy of current worksheet
    * @param [in] after If true, the new worksheet will be after this worksheet; 
    *                   otherwise, the new worksheet will be before this worksheet.
    * @return Return true if successful; otherwise, return false.
    * @note The new worksheet will be added after the current worksheet and will be activated.
    * @note Call ExcelWorkbook::GetActiveWorksheet() to get the new worksheet.
    */
    bool CopyWorksheet(bool after = true);

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
