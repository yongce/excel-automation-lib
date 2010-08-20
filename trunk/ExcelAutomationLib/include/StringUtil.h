/*!
* @file    StringUtil.h
* @brief   Header file for string utilities
* @date    2009-12-01
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef STRINGUTIL_H_GUID_07C67C12_CDD0_4E71_AB14_5800247ACD09
#define STRINGUTIL_H_GUID_07C67C12_CDD0_4E71_AB14_5800247ACD09


#include <string>
#include <sstream>
#include <fstream>
#include "LibDef.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


// macros for string literal
#ifdef _UNICODE

    #define ELtext(s)               L ## s
    typedef wchar_t                 ELchar;
    typedef std::wstring            ELstring;
    typedef std::wistringstream     EListringstream;
    typedef std::wostringstream     ELostringstream;
    typedef std::wifstream          ELifstream;
    typedef std::wofstream          ELofstream;
    typedef std::wostream           ELostream;

#else // non _UNICODE

    #define ELtext(s)               s
    typedef char                    ELchar;
    typedef std::string             ELstring;
    typedef std::istringstream      EListringstream;
    typedef std::ostringstream      ELostringstream;
    typedef std::ifstream           ELifstream;
    typedef std::ofstream           ELofstream;
    typedef std::ostream            ELostream;

#endif // _UNICODE


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END


#endif //STRINGUTIL_H_GUID_07C67C12_CDD0_4E71_AB14_5800247ACD09
