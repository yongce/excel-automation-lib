/*!
* @file    AtomicsUtil.h
* @brief   Header file for class AtomicsUtil.
* @date    2009-12-09
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef ATOMICSUTIL_H_GUID_BBB9EDC4_2368_4C83_9E93_CE3FF8B7BC98
#define ATOMICSUTIL_H_GUID_BBB9EDC4_2368_4C83_9E93_CE3FF8B7BC98


#include <windows.h>
#include "LibDef.h"


// namespace start
EXCEL_AUTOMATION_NAMESPACE_START


/*!
* @internal
* @brief Class AtomicsUtil is an utility class which provides some atomic operations on integer.
*        All members of AtomicsUtil are static members.
* @note AtomicsUtil is not intended and allowed to be instantiated.
*/
class AtomicsUtil
{
public:  // public types
    /*!
    * @brief Integer type for atomic operation.
    */
    typedef LONG Integer;

public:  // public interfaces
    /*!
    * @brief Increment an integer as an atomic operation.
    * @param pValue A pointer to the variable to be incremented.
    * @return The resulting incremented value.
    */
    static Integer Increment(Integer *pValue)
    {
        return ::InterlockedIncrement(pValue);
    }

    /*!
    * @brief Decrement an integer as an atomic operation.
    * @param pValue A pointer to the variable to be decremented.
    * @return The resulting decremented value.
    */
    static Integer Decrement(Integer *pValue)
    {
        return ::InterlockedDecrement(pValue);
    }

private:
    // Forbid instantiation
    AtomicsUtil();
};


// namespace end
EXCEL_AUTOMATION_NAMESPACE_END


#endif //ATOMICSUTIL_H_GUID_BBB9EDC4_2368_4C83_9E93_CE3FF8B7BC98
