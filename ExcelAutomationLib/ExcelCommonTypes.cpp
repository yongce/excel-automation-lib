/*!
* @file    ExcelCommonTypes.cpp
* @brief   Implementation file
* @date    2010-09-04 22:44:47
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#include "ExcelCommonTypes.h"

using namespace std;


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START



bool GetExcelConstant(ExcelHorizontalAlignment align, int &alignConstant)
{
    // Values of constants
    const int xlLeft = -4131;
    const int xlCenter = -4108;
    const int xlRight = -4152;

    switch (align)
    {
    case EHA_Left:
        alignConstant = xlLeft;
        break;

    case EHA_HCenter:
        alignConstant = xlCenter;
        break;

    case EHA_Right:
        alignConstant = xlRight;
        break;

    default:
        return false; // Unknow type
    }

    return true;
}


bool GetExcelConstant(ExcelVerticalAlignment align, int &alignConstant)
{
    // Values of constants
    const int xlTop = -4160;
    const int xlCenter = -4108;
    const int xlBottom = -4107;

    switch (align)
    {
    case EVA_Top:
        alignConstant = xlTop;
        break;

    case EVA_VCenter:
        alignConstant = xlCenter;
        break;

    case EVA_Bottom:
        alignConstant = xlBottom;
        break;

    default:
        return false; // Unknow type
    }

    return true;
}


int GuessFileFormatFromFilename(const ELstring &filename)
{
    int formatCode = 56;  // xlExcel8 (97-2003 format in Excel 2007-2010, xls), by default

    size_t lastDotPos = filename.find_last_of(ELtext('.'));

    if (lastDotPos != string::npos && lastDotPos + 1 < filename.length())
    {
        ELstring ext = filename.substr(lastDotPos + 1);
        if (ext == ELtext("xlsx"))
            formatCode = 51;  // xlOpenXMLWorkbook
    }

    return formatCode;
}



// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END
