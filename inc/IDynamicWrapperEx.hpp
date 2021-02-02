/**
* @file			IDynamicWrapperEx.hpp
* @date			01-10-2020
* @author		Paul Laine (@am0nsec)
* @version		1.0
* @brief		DynamicWrapperEX Automation declaration.
* @details
* @link         https://github.com/am0nsec/DynamicWrapperEx
* @copyright	This project has been released under the GNU Public License v3 license.
*/
#pragma once
#include <windows.h>
#include <unknwn.h>
#include <memory>
#include <vector>

#include "AutomationFactory.hpp"

#ifndef __IDYNAMICWRAPPEREX_H
#define __IDYNAMICWRAPPEREX_H

/**
 * @brief {F757F2EC-62D8-4BAE-8BE0-0A61CF36A541}
*/
static CONST GUID IID_IDynamicWrapperEx = { 0xf757f2ec , 0x62d8, 0x4bae , {0x8b, 0xe0, 0xa, 0x61, 0xcf, 0x36, 0xa5, 0x41} };

/**
 * @brief DynamicWrapperEx Automation Interface.
*/
class IDynamicWrapperEx : IDispatch {
public:
	/**
	 * @brief Constructor.
	*/
	IDynamicWrapperEx();

	/**
	 * @brief Destructor.
	*/
	virtual ~IDynamicWrapperEx();

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
	 * @brief Retrieves the number of type information interfaces that an object provides (either 0 or 1).
	 * @param pctinfo The number of type information interfaces provided by the object.
	 * @return Method not implemented.
	*/
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(
		_Out_ UINT* pctinfo
	);

	/**
	 * @brief Retrieves the type information for an object, which can then be used to get the type information for an interface.
	 * @param iTInfo The type information to return. Pass 0 to retrieve type information for the IDispatch implementation.
	 * @param lcid The locale identifier for the type information.
	 * @param ppTInfo The requested type information object.
	 * @return Whether the function executed successfully.
	*/
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(
		_In_  UINT        iTInfo,
		_In_  LCID        lcid,
		_Out_ ITypeInfo** ppTInfo
	);

	/**
	 * @brief Maps a single member and an optional set of argument names to a corresponding set of integer DISPIDs, which can be used on subsequent calls to Invoke.
	 * @param riid Reserved for future use. Must be IID_NULL.
	 * @param rgszNames The array of names to be mapped.
	 * @param cNames The count of the names to be mapped.
	 * @param lcid The locale context in which to interpret the names.
	 * @param rgDispId Caller-allocated array, each element of which contains an identifier (ID) corresponding to one of the names passed in the rgszNames array.
	 * @return Whether the function executed successfully.
	*/
	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(
		_In_  REFIID    riid,
		_In_  LPOLESTR* rgszNames,
		_In_  UINT      cNames,
		_In_  LCID      lcid,
		_Out_ DISPID*   rgDispId
	);

	/**
	 * @brief Provides access to properties and methods exposed by an object.
	 * @param dispIdMember Identifies the member.
	 * @param riid Reserved for future use. Must be IID_NULL.
	 * @param lcid The locale context in which to interpret arguments.
	 * @param wFlags Flags describing the context of the Invoke call.
	 * @param pDispParams Pointer to a DISPPARAMS structure containing an array of arguments, an array of argument DISPIDs for named arguments, and counts for the number of elements in the arrays.
	 * @param pVarResult Pointer to the location where the result is to be stored, or NULL if the caller expects no result.
	 * @param pExcepInfo Pointer to a structure that contains exception information.
	 * @param puArgErr The index within rgvarg of the first argument that has an error.
	 * @return Whether the function executed successfully.
	*/
	virtual HRESULT STDMETHODCALLTYPE Invoke(
		_In_  DISPID      dispIdMember,
		_In_  REFIID      riid,
		_In_  LCID        lcid,
		_In_  WORD        wFlags,
		_In_  DISPPARAMS* pDispParams,
		_Out_ VARIANT*    pVarResult,
		_Out_ EXCEPINFO*  pExcepInfo,
		_Out_ UINT*       puArgErr
	);

private:
	/**
	* @brief Number of reference to the object.
	*/
	DWORD m_dwReference{ 0 };

	/**
	 * @brief Unique pointer to an AutomationFactory class.
	*/
	std::unique_ptr<AutomationFactory> m_pAutomationFactory{ std::make_unique<AutomationFactory>() };
};

#endif // !__IDYNAMICWRAPPEREX_H
