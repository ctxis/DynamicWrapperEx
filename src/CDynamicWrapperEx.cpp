/**
* @file			CDynamicWrapperEx.cpp
* @date			01-10-2020
* @author		Paul Laine (@am0nsec)
* @version		1.0
* @brief		DynamicWrapperEx class factory definition.
* @details
* @link         https://github.com/am0nsec/DynamicWrapperEx
* @copyright	This project has been released under the GNU Public License v3 license.
*/
#include <windows.h>
#include <memory>

#include "IDynamicWrapperEx.hpp"
#include "CDynamicWrapperEx.hpp"

/**
 * @brief Constructor.
*/
CDynamicWrapperEx::CDynamicWrapperEx() { }

/**
 * @brief Destructor.
*/
CDynamicWrapperEx::~CDynamicWrapperEx() { }

/**
 * @brief Queries a COM object for a pointer to one of its interface.
 * @param riid A reference to the interface identifier (IID) of the interface being queried for.
 * @param ppvObject The address of a pointer to an interface with the IID specified in the riid parameter.
 * @return Whether an interface has been found.
*/
HRESULT STDMETHODCALLTYPE CDynamicWrapperEx::QueryInterface(
    _In_  REFIID  riid,
    _Out_ LPVOID* ppvObject
) {
    if (riid == IID_IUnknown || riid == IID_IClassFactory || riid == CLSID_CDynamicWrapperEx) {
        *ppvObject = this;
        this->AddRef();
        return S_OK;
    }

    *ppvObject = NULL;
    return E_NOINTERFACE;
}

/**
 * @brief  Increment the number of references.
 * @return Number of remaining references.
*/
ULONG STDMETHODCALLTYPE CDynamicWrapperEx::AddRef(VOID) {
    return InterlockedIncrement(&this->m_dwReference);
}

/**
 * @brief  Decrement the number of references.
 * @return Number of remaining references.
*/
ULONG STDMETHODCALLTYPE CDynamicWrapperEx::Release(VOID) {
    return InterlockedDecrement(&this->m_dwReference);
}

/**
 * @brief Creates an uninitialized object.
 * @param pUnkOuter Address to an IUknown interface if part of an aggregation.
 * @param riid A reference to the identifier of the interface to be used to communicate with the newly created object.
 * @param ppvObject The address of pointer variable that receives the interface pointer requested in riid.
 * @return Whether the object has been successfully created.
*/
HRESULT STDMETHODCALLTYPE CDynamicWrapperEx::CreateInstance(
    _In_  IUnknown* pUnkOuter,
    _In_  REFIID    riid,
    _Out_ LPVOID*   ppvObject
) {
    if (pUnkOuter != NULL)
        return CLASS_E_NOAGGREGATION;

    IDynamicWrapperEx* pIDynamicWrapperEx = new IDynamicWrapperEx();
    if (pIDynamicWrapperEx == nullptr)
        return E_OUTOFMEMORY;
    
    HRESULT hr = pIDynamicWrapperEx->QueryInterface(riid, ppvObject);
    this->m_pIDynamicWrapperEx = reinterpret_cast<IDynamicWrapperEx*>(ppvObject);
    pIDynamicWrapperEx->Release();
    return hr;
}

/**
 * @brief Locks an object application open in memory. This enables instances to be created more quickly.
 * @param fLock If TRUE, increments the lock count; if FALSE, decrements the lock count.
 * @return Whether the function executed successfully.
*/
HRESULT STDMETHODCALLTYPE CDynamicWrapperEx::LockServer(
    _In_ BOOL fLock
) {
    return ::CoLockObjectExternal(this, fLock, TRUE);
}