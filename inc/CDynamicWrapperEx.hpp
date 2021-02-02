/**
* @file			CDynamicWrapperEx.hpp
* @date			01-10-2020
* @author		Paul Laine (@am0nsec)
* @version		1.0
* @brief		DynamicWrapperEx class factory declaration.
* @details
* @link         https://github.com/am0nsec/DynamicWrapperEx
* @copyright	This project has been released under the GNU Public License v3 license.
*/
#pragma once
#include <windows.h>
#include <objbase.h>

#include "IDynamicWrapperEx.hpp"

#ifndef __CDYNAMICWRAPPEREX_HPP
#define __CDYNAMICWRAPPEREX_HPP

/**
 * @brief {1E2F6CDD-E721-4E94-885C-36C95D6A8CC2}
*/
static CONST GUID CLSID_CDynamicWrapperEx = { 0x1e2f6cdd, 0xe721, 0x4e94, {0x88, 0x5c, 0x36, 0xc9, 0x5d, 0x6a, 0x8c, 0xc2} };

/**
 * @brief DynamicWrapperEx class factory
*/
class CDynamicWrapperEx : IClassFactory {
public:
	/**
	 * @brief Constructor.
	*/
	CDynamicWrapperEx();

	/**
	 * @brief Destructor.
	*/
	virtual ~CDynamicWrapperEx();

	/**
	 * @brief Queries a COM object for a pointer to one of its interface.
	 * @param riid A reference to the interface identifier (IID) of the interface being queried for.
	 * @param ppvObject The address of a pointer to an interface with the IID specified in the riid parameter.
	 * @return Whether an interface has been found.
	*/
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(
		_In_  REFIID  riid, 
		_Out_ LPVOID* ppvObject
	);

	/**
	 * @brief  Increment the number of references.
	 * @return Number of remaining references.
	*/
	virtual ULONG STDMETHODCALLTYPE AddRef(VOID);

	/**
	 * @brief  Decrement the number of references.
	 * @return Number of remaining references.
	*/
	virtual ULONG STDMETHODCALLTYPE Release(VOID);

	/**
	 * @brief Creates an uninitialized object.
	 * @param pUnkOuter Address to an IUknown interface if part of an aggregation.
	 * @param riid A reference to the identifier of the interface to be used to communicate with the newly created object.
	 * @param ppvObject The address of pointer variable that receives the interface pointer requested in riid.
	 * @return Whether the object has been successfully created.
	*/
	virtual HRESULT STDMETHODCALLTYPE CreateInstance(
		_In_  IUnknown* pUnkOuter,
		_In_  REFIID    riid, 
		_Out_ LPVOID*   ppvObject
	);

	/**
	 * @brief Locks an object application open in memory. This enables instances to be created more quickly.
	 * @param fLock If TRUE, increments the lock count; if FALSE, decrements the lock count.
	 * @return Whether the function executed successfully.
	*/
	virtual HRESULT STDMETHODCALLTYPE LockServer(
		_In_ BOOL fLock
	);

private:
	/**
	 * @brief Number of reference to the object.
	*/
	DWORD m_dwReference{ 0 };

	/**
	 * @brief Unique pointer to an IDynamicWrapperEx COM Automation object.
	*/
	IDynamicWrapperEx* m_pIDynamicWrapperEx{ nullptr };
};

#endif // !__CDYNAMICWRAPPEREX_HPP
