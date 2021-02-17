/**
* @file			IDynamicWrapperEx.cpp
* @date			01-10-2020
* @author		Paul Laine (@am0nsec)
* @version		1.0
* @brief		DynamicWrapperEx Automation definition.
* @details
* @link         https://github.com/am0nsec/DynamicWrapperEx
* @copyright	This project has been released under the GNU Public License v3 license.
*/
#include <windows.h>
#include <unknwn.h>
#include <memory>

#include "IDynamicWrapperEx.hpp"
#include "AutomationFactory.hpp"
#include "Util.hpp"

/**
 * @brief Constructor.
*/
IDynamicWrapperEx::IDynamicWrapperEx() {
	this->m_pAutomationFactory->m_aDispatchTable.push_back({ 0, L"DwRegister"});
	this->m_pAutomationFactory->m_aDispatchTable.push_back({ 1, L"WriteByte" });

	this->m_pAutomationFactory->m_dwInternalMethods = this->m_pAutomationFactory->m_aDispatchTable.size();
}

/**
 * @brief Destructor.
*/
IDynamicWrapperEx::~IDynamicWrapperEx() { }

/**
*@brief Queries a COM object for a pointer to one of its interface.
* @param riid A reference to the interface identifier(IID) of the interface being queried for.
* @param ppvObject The address of a pointer to an interface with the IID specified in the riid parameter.
* @return Whether an interface has been found.
*/
HRESULT STDMETHODCALLTYPE IDynamicWrapperEx::QueryInterface(
	_In_  REFIID  riid,
	_Out_ LPVOID* ppvObject
) {
	if (IsEqualGUID(riid, IID_IDynamicWrapperEx) || IsEqualGUID(riid, IID_IDispatch) || IsEqualGUID(riid, IID_IUnknown)) {
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
ULONG STDMETHODCALLTYPE IDynamicWrapperEx::AddRef(VOID) {
	return InterlockedIncrement(&this->m_dwReference);
}

/**
 * @brief  Decrement the number of references.
 * @return Number of remaining references.
*/
ULONG STDMETHODCALLTYPE IDynamicWrapperEx::Release(VOID) {
	ULONG ulReference = InterlockedDecrement(&this->m_dwReference);
	return ulReference;
}

/**
 * @brief Retrieves the number of type information interfaces that an object provides (either 0 or 1).
 * @param pctinfo The number of type information interfaces provided by the object.
 * @return Method not implemented.
*/
HRESULT STDMETHODCALLTYPE IDynamicWrapperEx::GetTypeInfoCount(
	_Out_ UINT* pctinfo
) {
	*pctinfo = 1;
	return E_NOTIMPL;
}

/**
 * @brief Retrieves the type information for an object, which can then be used to get the type information for an interface.
 * @param iTInfo The type information to return. Pass 0 to retrieve type information for the IDispatch implementation.
 * @param lcid The locale identifier for the type information.
 * @param ppTInfo The requested type information object.
 * @return Whether the function executed successfully.
*/
HRESULT STDMETHODCALLTYPE IDynamicWrapperEx::GetTypeInfo(
	_In_  UINT iTInfo,
	_In_  LCID lcid,
	_Out_ ITypeInfo** ppTInfo
) {
	if (ppTInfo != NULL)
		*ppTInfo = NULL;
	return S_OK;
}

/**
 * @brief Maps a single member and an optional set of argument names to a corresponding set of integer DISPIDs, which can be used on subsequent calls to Invoke.
 * @param riid Reserved for future use. Must be IID_NULL.
 * @param rgszNames The array of names to be mapped.
 * @param cNames The count of the names to be mapped.
 * @param lcid The locale context in which to interpret the names.
 * @param rgDispId Caller-allocated array, each element of which contains an identifier (ID) corresponding to one of the names passed in the rgszNames array.
 * @return Whether the function executed successfully.
*/
HRESULT STDMETHODCALLTYPE IDynamicWrapperEx::GetIDsOfNames(
	_In_  REFIID riid,
	_In_  LPOLESTR* rgszNames, 
	_In_  UINT cNames, 
	_In_  LCID lcid,
	_Out_ DISPID* rgDispId
) {
	if (riid != IID_NULL)
		return DISP_E_UNKNOWNINTERFACE;

	for (auto& elem : this->m_pAutomationFactory->m_aDispatchTable) {
		if (::wcscmp(elem.wszName, *rgszNames) == 0) {
			*rgDispId = elem.lDispId;
			return S_OK;
		}
	}

	*rgDispId = 0;
	return E_FAIL;
}

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
HRESULT STDMETHODCALLTYPE IDynamicWrapperEx::Invoke(
	_In_  DISPID dispIdMember,
	_In_  REFIID riid, 
	_In_  LCID lcid,
	_In_  WORD wFlags,
	_In_  DISPPARAMS* pDispParams,
	_Out_ VARIANT* pVarResult,
	_Out_ EXCEPINFO* pExcepInfo,
	_Out_ UINT* puArgErr
) {
	if (riid != IID_NULL)
		return DISP_E_UNKNOWNINTERFACE;

	if ((wFlags & DISPATCH_METHOD) != DISPATCH_METHOD)
		return E_FAIL;

	// Non-dynamic methods
	if (dispIdMember == 0)
		return this->m_pAutomationFactory->Register(pDispParams, pVarResult);
	else if (dispIdMember == 1)
		return Util::WriteByte(pDispParams, pVarResult);

	// Execute dynamic method
	HRESULT hr = this->m_pAutomationFactory->m_aDynamicMethods[dispIdMember - this->m_pAutomationFactory->m_dwInternalMethods]->Invoke(pDispParams, pVarResult);
	return hr;
}
