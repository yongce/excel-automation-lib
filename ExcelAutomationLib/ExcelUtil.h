/*!
* @file    ExcelUtil.h
* @brief   Header file for class ExcelUtil
* @date    2010-09-07 15:01:24
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef EXCELUTIL_H_GUID_FBB6D52B_54B3_4DC5_A5B5_E746A3A94DB4
#define EXCELUTIL_H_GUID_FBB6D52B_54B3_4DC5_A5B5_E746A3A94DB4


#include "LibDef.h"
#include "ExcelCommonTypes.h"
#include "StringUtil.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


class ExcelUtil
{
public:
    static bool GetExcelConstant(ExcelHorizontalAlignment align, int &alignConstant);
    static bool GetExcelConstant(ExcelVerticalAlignment align, int &alignConstant);

    //  Guess file format from file name.
    static int GuessFileFormatFromFilename(const ELstring &filename);
};


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END


#endif //EXCELUTIL_H_GUID_FBB6D52B_54B3_4DC5_A5B5_E746A3A94DB4

