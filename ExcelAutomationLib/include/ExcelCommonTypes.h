/*!
* @file    ExcelCommonTypes.h
* @brief   Header file
* @date    2010-09-04 21:33:10
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef EXCELCOMMONTYPES_H_GUID_0A4BE828_D69E_4A79_A0C1_FA633D6C6035
#define EXCELCOMMONTYPES_H_GUID_0A4BE828_D69E_4A79_A0C1_FA633D6C6035


#include "LibDef.h"
#include "StringUtil.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


// Constants for Horizontal Alignment
enum ExcelHorizontalAlignment
{
    EHA_Left = 1,             // Left alignment
    EHA_HCenter = 2,          // Horizontal center alignment: 
    EHA_Right = 3,            // Right alignment
};

bool GetExcelConstant(ExcelHorizontalAlignment align, int &alignConstant);

// Constants for Vertical Alignment
enum ExcelVerticalAlignment
{
    EVA_Top = 1,              // Top alignment
    EVA_VCenter = 2,          // Vertical center alignment
    EVA_Bottom = 3,           // Bottom alignment
};

bool GetExcelConstant(ExcelVerticalAlignment align, int &alignConstant);


//  Guess file format from file name. Default is 
int GuessFileFormatFromFilename(const ELstring &filename);


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END


#endif //EXCELCOMMONTYPES_H_GUID_0A4BE828_D69E_4A79_A0C1_FA633D6C6035

