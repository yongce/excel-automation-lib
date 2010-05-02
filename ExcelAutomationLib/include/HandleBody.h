/*!
* @file    HandleBody.h
* @brief   Header file for Handle/Body pattern
* @date    2009-12-09
* @author  Tu Yongce <tuyongce@gmail.com>
* @version $Id$
*/


#ifndef HANDLEBODY_H_GUID_2A95B06A_A28C_4955_B25B_6D7AB9A8ABA7
#define HANDLEBODY_H_GUID_2A95B06A_A28C_4955_B25B_6D7AB9A8ABA7


/*!
* @file 
* This file implements the Handle/Body pattern. 
* Class HandleBase is the base class for Handle and class BodyBase is the base class for Body. @n
* At the end of this file, one example of implementing Handle/Body pattern is shown.
*/


#include <cassert>
#include "LibDef.h"
#include "AtomicsUtil.h"


// namespace start
EXCEL_AUTOMATION_NAMESPACE_START


/*!
* @internal
* @brief Base class for "Body" in the Handle/Body pattern.
*/
class EXCEL_AUTOMATION_DLL_API BodyBase
{
    friend class HandleBase;

protected:
    BodyBase(): m_refCount(0) { }
    virtual ~BodyBase() { }

    AtomicsUtil::Integer m_refCount;
};


/*!
* @internal
* @brief HandleNullPtr is a dummy class with no implementation, 
*        which presents a null pointer type.
*/
class HandleNullPtr;


/*!
* @internal
* @brief nullHandle represents a null pointer of the "Handle" class
*/
const HandleNullPtr * const nullHandle = 0;


/*!
* @internal
* @brief Base class for "Handle" in the Handle/Body pattern.
*/
class EXCEL_AUTOMATION_DLL_API HandleBase
{
public:
    bool IsNull() const
    {
        return m_pBody == 0;
    }

    // <start> Support comparing with null pointer
    bool operator == (const HandleNullPtr *)
    {
        return m_pBody == 0;
    }

    bool operator != (const HandleNullPtr *)
    {
        return m_pBody != 0;
    }
    // <end> Support comparing with null pointer

protected: 
    HandleBase(BodyBase *pBody = 0): m_pBody(pBody)
    {
        AddRef();
    }

    HandleBase(const HandleBase &other): m_pBody(other.m_pBody)
    {
        AddRef();
    }

    HandleBase& operator = (const HandleBase &rhs)
    {
        if (&rhs != this)
        {
            ReleaseRef();
            m_pBody = rhs.m_pBody;
            AddRef();
        }

        return *this;
    }

    // <start> Support comparing with null pointer
    HandleBase& operator = (const HandleNullPtr *ptr)
    {
        assert(ptr == nullHandle);
        (ptr);              // eliminate "unreferenced formal parameter" warning
        ReleaseRef();
        return *this;
    }
    // <end> Support comparing with null pointer

    virtual ~HandleBase()
    {
        ReleaseRef();
    }

    BodyBase& Body() const
    {
        assert(m_pBody);
        return *m_pBody;
    }

private:
    void AddRef()
    {
        if (m_pBody)
        {
            AtomicsUtil::Increment(&(m_pBody->m_refCount));
        }
    }

    void ReleaseRef()
    {
        if (m_pBody)
        {
            if (AtomicsUtil::Decrement(&(m_pBody->m_refCount)) == 0)
            {
                delete m_pBody;
            }
            m_pBody = 0;
        }
    }

protected:
    BodyBase *m_pBody;
};


/*!
* @internal
* @brief A handle class as a observer for the body class.
* @tparam THandle A "Handle" class which should be a class derived from HandleBase.
* @tparam TBody A "Body" class which should be a class derived from BodyBase.
* @note If FriendHandle<THanlde, TBody> is used, it must be declared as a friend class of THandle.
*/
template <class THandle, class TBody>
class FriendHandle
{
public:
    FriendHandle(): m_pBody(0) { }

    FriendHandle& operator = (const THandle &rhs)
    {
        m_pBody = dynamic_cast<TBody*>(rhs.m_pBody);
        return *this;
    }

    THandle GetHandle() const
    {
        THanlde(m_pBody);
    }

    // Support comparing with null pointer
    bool operator == (const HandleNullPtr *ptr) const
    {
        return m_pBody == 0;
    }

    bool operator != (const HandleNullPtr *ptr) const
    {
        return m_pBody != 0;
    }

private:
    TBody *m_pBody;
};



/////////////////////////////////////////////////////////////////////////////////////////////////////////////
// An example to show how to use the Handle/Body pattern.

#if 0

// MyClass.h

class MyClassImpl;

class MyClass : public HandleBase
{
    friend class MyClassImpl;

public:
    MyClass() { } // if null handle is allowed

    MyClass(X x, Y y, Z z);  // Normal constructor

    R NormalMethod(X x, Y y, Z z); // Normal method

    // Following is the implementation of Handle/Body pattern

public:
    // if null handle is allowed, then add the following function
    MyHandle& operator = (const HandleNullPtr *ptr)
    {
        HandleBase::operator =(ptr);
        return *this;
    }

private:
    // the following code may be templated
    MyClass(MyClassImpl *impl): HandleBase(impl) { }

    MyClassImpl& Body() const
    {
        return (MyClassImpl&)HandleBase::Body();
    }

};


// MyClass.cpp

class MyClassImpl : public BodyBase
{
public:
    // methods corresponding to My class
    MyClassImpl(X x, Y y, Z z) { ...}
    R NormalMethod(X x, Y y, Z z) { ... }

    // Following are implementation of Handle/Body pattern 
    MyClass SomeMethod(...)
    {
        //...
        return MyClass(this);
    }


};


MyClass::MyClass(X x, Y y, Z z): HandleBase(new MyClssImpl(x, y, z))
{
}

R MyClass::NormalMethod(X x, Y y, Z z)
{
    return Body().NormalMethod(x, y, z);
}


#endif // #if 0


// namespace end
EXCEL_AUTOMATION_NAMESPACE_END


#endif //HANDLEBODY_H_GUID_2A95B06A_A28C_4955_B25B_6D7AB9A8ABA7
