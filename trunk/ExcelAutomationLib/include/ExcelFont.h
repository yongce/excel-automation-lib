/*!
* @file    ExcelFont.h
* @brief   Header file for class ExcelFont
* @date    2010-09-01 15:12:42
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef EXCELFONT_H_GUID_76A3C90B_F8F4_49A5_9F56_9CA86958D756
#define EXCELFONT_H_GUID_76A3C90B_F8F4_49A5_9F56_9CA86958D756


#include "LibDef.h"
#include "HandleBody.h"
#include "StringUtil.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


/*!
* @brief Class ExcelFont represents the concept "Font" in Excel.
* @note ExcelFont/ExcelFontImpl is an implementation of the "Handle/Body" pattern.
*/
class EXCEL_AUTOMATION_DLL_API ExcelFont : public HandleBase
{
public:
    /*!
    * Default constructor
    */ // Doc is needed by Doxygen
    ExcelFont() : HandleBase(0) { }

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
    friend class ExcelRangeImpl;  // which will call the following ctor
    friend class ExcelCellImpl;   // which will call the following ctor
    ExcelFont(IDispatch *pFont);

private:
    // <begin> Handle/Body pattern implementation
    friend class ExcelFontImpl;
    ExcelFont(ExcelFontImpl *impl);
    ExcelFontImpl& Body() const;
    // <end> Handle/Body pattern implementation
};


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END


#endif //EXCELFONT_H_GUID_76A3C90B_F8F4_49A5_9F56_9CA86958D756

