/*!
* @file    Noncopyable.h
* @brief   Header file for class Noncopyable
* @date    2009-12-01
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef NONCOPYABLE_H_GUID_77EEEF35_EF3F_4B71_AA1E_5AECDA38DDAB
#define NONCOPYABLE_H_GUID_77EEEF35_EF3F_4B71_AA1E_5AECDA38DDAB


#include "LibDef.h"


// <begin> namespace
EXCEL_AUTOMATION_NAMESPACE_START


/*!
* @brief    An utility class used to forbid copy ctor & copy assignment.
* @details  Any class can inherit this class to forbid its own copy ctor & copy assignment.
*/
class Noncopyable
{
protected:
    Noncopyable() {}
    ~Noncopyable() {}

private:
    // Forbid copy ctor & copy assignment. No need to implement them.
    Noncopyable(const Noncopyable &);
    Noncopyable& operator = (const Noncopyable &);
};


// <end> namespace
EXCEL_AUTOMATION_NAMESPACE_END


#endif //NONCOPYABLE_H_GUID_77EEEF35_EF3F_4B71_AA1E_5AECDA38DDAB
