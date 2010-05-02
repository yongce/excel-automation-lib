/*!
* @file    ExcelCell.h
* @brief   Header file for class ExcelCell
* @date    2009-12-31
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef EXCELCELL_H_GUID_F7A9BE95_9657_4811_94CF_CB56C002DDF0
#define EXCELCELL_H_GUID_F7A9BE95_9657_4811_94CF_CB56C002DDF0


#include "LibDef.h"
#include "HandleBody.h"
#include "StringUtil.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


/*!
* @brief Class ExcelCell represents the concept "Cell" in Excel.
* @note ExcelCell/ExcelCellImpl is an implementation of the "Handle/Body" pattern.
* @note In fact, there is no such a "Cell" object in Excel Object Model. "Cell" is a special "Range". 
*       ExcelCell is provided just for convenience.
*/
class EXCEL_AUTOMATION_DLL_API ExcelCell : public HandleBase
{
public:
    /*!
    * Default constructor
    */ // Doc is needed by Doxygen
    ExcelCell() : HandleBase(0) { }

    bool GetValue(ELstring &value);
    bool SetValue(const ELstring &value);
    bool SetValue(int value);
    bool SetValue(double value);

private:
    friend class ExcelWorksheetImpl;  // which will call the following ctor
    ExcelCell(IDispatch *pCell, ELchar column, int row);

private:
    // <begin> Handle/Body pattern implementation
    friend class ExcelCellImpl;
    ExcelCell(ExcelCellImpl *impl);
    ExcelCellImpl& Body() const;
    // <end> Handle/Body pattern implementation
};


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END


#endif //EXCELCELL_H_GUID_F7A9BE95_9657_4811_94CF_CB56C002DDF0
