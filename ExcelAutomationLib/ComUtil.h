/*!
* @file    ComUtil.h
* @brief   Header file for class ComUtil
* @date    2009-12-09
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef COMUTIL_H_GUID_79FBF1F4_971A_489D_AD88_92DE25791DC9
#define COMUTIL_H_GUID_79FBF1F4_971A_489D_AD88_92DE25791DC9


#include <windows.h>
#include "LibDef.h"
#include "StringUtil.h"


// namespace start
EXCEL_AUTOMATION_NAMESPACE_START


/*!
* @brief Class ComUtil is an utility class which provides some wrapper functions for COM operation. 
*        All the members of ComUtil are static member.
* @note ComUtil is not intended and allowed to be instantiated.
*/
class ComUtil
{
public:
    /*!
    * @brief ComUtil::Invoke is a wrapper of IDispatch::Invoke(), which is provided to simplify our work.
    * @param [in] pDisp Pointer to IDispatch. Must not be NULL.
    * @param [in] type Three values are allowed: DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT
    * @param [in] name Name of the method or property involved.
    * @param [out] pResult Pointer to a variant which holds the results. Can be NULL.
    * @param [in] argc Number of the variant parameters. Must be equal or greater than 0.
    * @param [in] ... The variant parameters list. Their types must be VARIANT.
    * @return Any value which can be returned by IDispatch::GetIDsOfNames() or IDispatch::Invoke().
    */
    static HRESULT Invoke(IDispatch *pDisp, WORD type, LPOLESTR name, VARIANT *pResult, int argc, ...);

    /*!
    * @brief Get an element of a two-dimensional SAFEARRAY. 
    *        ComUtil::GetSafeArrayElementDim2() is a wrapper of ::SafeArrayGetElements().
    * @param [in] psa Pointer to an SAFEARRAY object which should be a two-dimensional array. 
    *                 Must not be NULL. Element type of the SAFEARRAY object must be VARIANT.
    * @param [in] dim1 Index of the first dimension of the array.
    * @param [in] dim2 Index of the second dimension of the array.
    * @param [out] pResult Pointer to a VARIANT object to save the result. Must not be NULL.
    * @return Any value which can be returned by ::SafeArrayGetElements().
    */
    static HRESULT GetSafeArrayElementDim2(SAFEARRAY *psa, long dim1, long dim2, VARIANT *pResult);

    /*!
    * @brief Set value of type string for the element in a two-dimensional SAFEARRAY. 
    *        ComUtil::PutSafeArrayElementDim2() is a wrapper of ::SafeArrayPutElement().
    * @param [in,out] psa Pointer to an SAFEARRAY object which should be a two-dimensional array. 
    *                     Must not be NULL. Element type of the SAFEARRAY object must be VARIANT.
    * @param [in] dim1 Index of the first dimension of the array.
    * @param [in] dim2 Index of the second dimension of the array.
    * @param [in] value An OLE string which will be saved into the array.
    * @return Any value which can be returned by ::SafeArrayPutElement().
    */
    static HRESULT PutSafeArrayElementDim2(SAFEARRAY *psa, long dim1, long dim2, const OLECHAR *value);

    /*!
    * @brief Set value of type int for the element in a two-dimensional SAFEARRAY.
    * @see HRESULT ComUtil::PutSafeArrayElementDim2(SAFEARRAY *psa, long dim1, long dim2, const OLECHAR *value)
    */
    static HRESULT PutSafeArrayElementDim2(SAFEARRAY *psa, long dim1, long dim2, int value);

    /*!
    * @brief Set value of type doulbe for the element in a two-dimensional SAFEARRAY.
    * @see HRESULT ComUtil::PutSafeArrayElementDim2(SAFEARRAY *psa, long dim1, long dim2, const OLECHAR *value)
    */
    static HRESULT PutSafeArrayElementDim2(SAFEARRAY *psa, long dim1, long dim2, double value);


    /*!
    * @brief Encode values in a two-dimensional SAFEARRAY object into string form.
    * @param [in] psa Pointer to an SAFEARRAY object which should be a two-dimensional array. 
    *                 Must not be NULL. Element type of the SAFEARRAY object must be VARIANT.
    * @param [out] encodedStr The corresponding encoded string of the array.
    * @return Any value which can be returned by ::SafeArrayGetLBound(), ::SafeArrayGetUBound(), 
    *         ::VariantChangeType() or ComUtil::GetSafeArrayElementDim2().
    * @note Encoding format: @n
    *         EncodingString := <number of rows>#<number of columns>#{Values}            @n
    *         Values := {RowValues}{RowValues}...{RowValues}              --> All rows in the two-dimentional array    @n
    *         RowValues := {ColumnValue}{ColumnValue}...{ColumnValue}     --> All columns of in one row     @n
    *         ColumnValue := <number of characters>#<the character string of the value>          @n
    *       Example: @n
    *         For the two-dimensional array (2*5) @n
    *           { {abc, de, fghi, 3, 5235}, {23, 5353, 3253, 32} }, @n
    *         The correpsonding encoded string is 2#5#3#abc2#de4#fghi1#34#52352#234#53530#4#32532#32.    @n
    * @todo Add type info into the encoded string.
    */
    static HRESULT EncodeSafeArrayDim2(SAFEARRAY *psa, ELstring &encodedStr);

    /*!
    * @brief Decode the data from the encoded string and create an SAFEARRAY to store the data.
    * @param [in] data The encoded string of a two dimensional array.
    * @return An SAFEARRAY with decoded data from encoded string. 
    *         If failed to decode the <em>data</em>, NULL will be returned.
    * @note The encoding format of @e data must be the one specified in ComUtil::EncodeSafeArrayDim2().
    */
    static SAFEARRAY* DecodeSafeArrayDim2(const ELchar *data);

private:
    // Forbid instantiation
    ComUtil();
};


// namespace end
EXCEL_AUTOMATION_NAMESPACE_END


#endif //COMUTIL_H_GUID_79FBF1F4_971A_489D_AD88_92DE25791DC9
