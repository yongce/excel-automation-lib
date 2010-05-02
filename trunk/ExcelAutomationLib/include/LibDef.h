/*!
* @file    LibDef.h
* @brief   Header file for macro definitions
* @date    2009-12-01
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef LIBDEF_H_GUID_03294A0C_978E_472D_853A_7208AE9990FB
#define LIBDEF_H_GUID_03294A0C_978E_472D_853A_7208AE9990FB


/*!
* @brief All things provided by ExcelAutomationLib are in namespace ExcelAutomation.
*/
namespace ExcelAutomation {} // this definition only for generating doc


// macros defined for convenience
#define EXCEL_AUTOMATION_NAMESPACE_START namespace ExcelAutomation {
#define EXCEL_AUTOMATION_NAMESPACE_END   }


// Dll API
#ifdef EXCEL_AUTOMATION_LIB_BUILD
#   define EXCEL_AUTOMATION_DLL_API __declspec(dllexport)
#else
#   define EXCEL_AUTOMATION_DLL_API __declspec(dllimport)
#endif



#endif //LIBDEF_H_GUID_03294A0C_978E_472D_853A_7208AE9990FB
