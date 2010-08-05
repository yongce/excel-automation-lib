/*!
* @file    ComUtil.cpp
* @brief   Implementation file for class ComUtil
* @date    2009-12-09
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#include <tchar.h>
#include <cassert>
#include <sstream>
#include <iomanip>
#include "ComUtil.h"


// namespace start
EXCEL_AUTOMATION_NAMESPACE_START


/*
* @brief ComUtil::Invoke is a wrapper of IDispatch::Invoke(), which is provided to simplify our work.
* @param [in] pDisp Pointer to IDispatch. Must not be NULL.
* @param [in] type Three values are allowed: DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT
* @param [in] name Name of the method or property involved.
* @param [out] pResult Pointer to a variant which holds the results. Can be NULL.
* @param [in] argc Number of the variant parameters. Must be equal or greater than 0.
* @param [in] ... The variant parameters list. Their types must be VARIANT.
* @return Any value which can be returned by IDispatch::GetIDsOfNames() or IDispatch::Invoke().
*/
HRESULT ComUtil::Invoke(IDispatch *pDisp, WORD type, LPOLESTR name, VARIANT *pResult, int argc, ...)
{
    assert(pDisp);
    assert(argc >= 0);

    // setup the parameters
    va_list marker;
    va_start(marker, argc);

    VARIANT *pArgs = new VARIANT[argc + 1];
    for (int i = 0; i < argc; ++i)
    {
        pArgs[i] = va_arg(marker, VARIANT);
    }

    va_end(marker);

    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;

    dp.cArgs = argc;
    dp.rgvarg = pArgs;

    if (type & DISPATCH_PROPERTYPUT)
	{
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    // get ID of the name
    DISPID dispID;
    HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_SYSTEM_DEFAULT, &dispID);

    // do the invocation
    if (SUCCEEDED(hr))
    {
        hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT,	type, &dp, pResult, NULL, NULL);
    }

    delete[] pArgs;

    return hr;
}


/*
* @brief Get an element of a two-dimensional SAFEARRAY. 
*        ComUtil::GetSafeArrayElementDim2() is a wrapper of ::SafeArrayGetElements().
* @param [in] psa Pointer to an SAFEARRAY object which should be a two-dimensional array. 
*                 Must not be NULL. Element type of the SAFEARRAY object must be VARIANT.
* @param [in] dim1 Index of the first dimension of the array.
* @param [in] dim2 Index of the second dimension of the array.
* @param [out] pResult Pointer to a VARIANT object to save the result. Must not be NULL.
* @return Any value which can be returned by ::SafeArrayGetElements().
*/
HRESULT ComUtil::GetSafeArrayElementDim2(SAFEARRAY *psa, long dim1, long dim2, VARIANT *pResult)
{
    assert(psa && pResult);
    assert(::SafeArrayGetDim(psa) == 2);

	long indices[2];
	indices[0] = dim1;
	indices[1] = dim2;

	VariantInit(pResult);
	
    HRESULT hr = ::SafeArrayGetElement(psa, indices, pResult);

	return hr;
}


/*
* @brief Set value of type string for the element in a two-dimensional SAFEARRAY. 
*        ComUtil::PutSafeArrayElementDim2() is a wrapper of ::SafeArrayPutElement().
* @param [in,out] psa Pointer to an SAFEARRAY object which should be a two-dimensional array. 
*                     Must not be NULL. Element type of the SAFEARRAY object must be VARIANT.
* @param [in] dim1 Index of the first dimension of the array.
* @param [in] dim2 Index of the second dimension of the array.
* @param [in] value An OLE string which will be saved into the array.
* @return Any value which can be returned by ::SafeArrayPutElement().
*/
HRESULT ComUtil::PutSafeArrayElementDim2(SAFEARRAY *psa, long dim1, long dim2, const OLECHAR *value)
{
    assert(psa);
    assert(::SafeArrayGetDim(psa) == 2);
	
	long indices[2];
	indices[0] = dim1;
	indices[1] = dim2;
	
	VARIANT param;
	param.vt = VT_BSTR;
	param.bstrVal = ::SysAllocString(value);
	
    HRESULT hr = ::SafeArrayPutElement(psa, indices, &param);

	VariantClear(&param);

	return hr;
}


/*
* @brief Set value of type int for the element in a two-dimensional SAFEARRAY.
* @see HRESULT ComUtil::PutSafeArrayElementDim2(SAFEARRAY *psa, long dim1, long dim2, const OLECHAR *value)
*/
HRESULT ComUtil::PutSafeArrayElementDim2(SAFEARRAY *psa, long dim1, long dim2, int value)
{
    assert(psa);
    assert(::SafeArrayGetDim(psa) == 2);
	
	long indices[2];
	indices[0] = dim1;
	indices[1] = dim2;
	
	VARIANT param;
	param.vt = VT_INT;
	param.intVal = value;
	
    HRESULT hr = ::SafeArrayPutElement(psa, indices, &param);

	return hr;
}


/*
* @brief Set value of type doulbe for the element in a two-dimensional SAFEARRAY.
* @see HRESULT ComUtil::PutSafeArrayElementDim2(SAFEARRAY *psa, long dim1, long dim2, const OLECHAR *value)
*/
HRESULT ComUtil::PutSafeArrayElementDim2(SAFEARRAY *psa, long dim1, long dim2, double value)
{
    assert(psa);
    assert(::SafeArrayGetDim(psa) == 2);
	
	long indices[2];
	indices[0] = dim1;
	indices[1] = dim2;
	
	VARIANT param;
	param.vt = VT_R8;
	param.dblVal = value;
	
    HRESULT hr = ::SafeArrayPutElement(psa, indices, &param);

	return hr;
}


/*
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
HRESULT ComUtil::EncodeSafeArrayDim2(SAFEARRAY *psa, ELstring &encodedStr)
{
    assert(psa);
    assert(::SafeArrayGetDim(psa) == 2);

    LONG rowFrom = 0;
    LONG rowTo = 0;
    LONG columnFrom = 0;
    LONG columnTo = 0;

    HRESULT hr;
    hr = ::SafeArrayGetLBound(psa, 1, &rowFrom);
    if (FAILED(hr))
        return hr;

    hr = ::SafeArrayGetUBound(psa, 1, &rowTo);
    if (FAILED(hr))
        return hr;

    hr = ::SafeArrayGetLBound(psa, 2, &columnFrom);
    if (FAILED(hr))
        return hr;

    hr = ::SafeArrayGetUBound(psa, 2, &columnTo);
    if (FAILED(hr))
        return hr;


    ELostringstream oss;

    // Encoding format: <row>#<column>#
    oss << (rowTo - rowFrom + 1) << ELtext('#') << (columnTo - columnFrom + 1) << ELtext('#');

    for (LONG i = rowFrom; i <= rowTo; ++i)
    {
        for (LONG j = columnFrom; j <= columnTo; ++j)
        {
            VARIANT var;
            hr = GetSafeArrayElementDim2(psa, i, j, &var);
            if (FAILED(hr))
                return hr;

            // convert the VARIANT object into a string value
            hr = ::VariantChangeType(&var, &var, VARIANT_NOUSEROVERRIDE, VT_BSTR);
            if (FAILED(hr))
                return hr;

            ELstring str(var.bstrVal);

            // Encoding format: <number of characters>#<characters>
            oss << str.length() << ELtext('#') << str;
        }
    }

    // return the encoded string
    encodedStr = oss.str();

    return S_OK;
}


/*
* @brief Decode the data from the encoded string and create an SAFEARRAY to store the data.
* @param [in] data The encoded string of a two dimensional array.
* @return An SAFEARRAY with decoded data from encoded string. 
*         If failed to decode the <em>data</em>, NULL will be returned.
*/
SAFEARRAY* ComUtil::DecodeSafeArrayDim2(const ELchar *data)
{
    assert(data);

    EListringstream iss(data);
    iss >> std::noskipws;

    int row = 0;
    int column = 0;
    ELchar dumb;

    // Encoding format: <row>#<column>#
    iss >> row >> dumb;             // <row>#
    assert(dumb == ELtext('#'));    // Whether the delimiter is '#', it's not important. 
                                    // So we just place an assert statement here.
    iss >> column >> dumb;          // <column>#
    assert(dumb == ELtext('#'));

    if (row <= 0 || column <= 0)
        return NULL;   // no data or dirty data

    SAFEARRAYBOUND sab[2];
    sab[0].lLbound = 1;
    sab[0].cElements = row;
    sab[1].lLbound = 1;
    sab[1].cElements = column;

    SAFEARRAY *psa = ::SafeArrayCreate(VT_VARIANT, 2, sab);
    if (!psa)
        return NULL;   // failed to create an array

    bool validState = true;  // flag indicating whether the data is well formed

    for (int i = 1; validState && i <= row; ++i)
    {
        for (int j = 1; validState && j <= column; ++j)
        {
            // Encoding format: <number of characters>#<characters>
            int count = 0;
            iss >> count >> dumb;
            assert(dumb == ELtext('#'));

            validState = iss.good() && (count >= 0);
            
            ELstring value;
            for (int k = 0; k < count; ++k)
            {
                ELchar ch;
                iss >> ch;
                value.push_back(ch);
            }

            validState = validState && iss.good();

            if (validState)
                PutSafeArrayElementDim2(psa, i, j, value.c_str());
        }
    }

    if (!validState) {
        ::SafeArrayDestroy(psa);
        psa = NULL;
    }

    return psa;
}



// namespace end
EXCEL_AUTOMATION_NAMESPACE_END
