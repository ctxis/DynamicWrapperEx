/**
* @file			DynamicMethod.hpp
* @date			01-10-2020
* @author		Paul Laine (@am0nsec)
* @version		1.0
* @brief		Dynamic method declaration.
* @details
* @link         https://github.com/am0nsec/DynamicWrapperEx
* @copyright	This project has been released under the GNU Public License v3 license.
*/
#pragma once
#include <windows.h>

#include "Types.hpp"

#ifndef __DYNAMICMETHOD_HPP
#define __DYNAMICMETHOD_HPP

#define ARGUMENT_STD 0x00000000 /* Standard data */
#define ARGUMENT_FLT 0x00000002 /* Floating point data */
#define RETURN_STD   0x00000000 /* Standard data */
#define RETURN_FLT   0x00000002 /* Floating point data */

class DynamicMethod {
public:
	/**
	 * @brief Constructor.
	 * @param dwDispatchId The dispatch ID that has been associated to this dynamic method.
	 * @param bstrFunctionName The name of the function to execute.
	 * @param lpFunction The address of the function to execute.
	*/
	DynamicMethod(
		_In_ DWORD  dwDispatchId,
		_In_ BSTR   bstrFunctionName,
		_In_ LPVOID lpFunction
	);

	/**
	 * @brief Destructor.
	 */
	~DynamicMethod();

	/**
	 * @brief Dynamically execute the function associated to this dynamic method.
	 * @param pDispParams List of parameters provided by the client.
	 * @param pVarResult Pointer to the location where the result is to be stored, or NULL if the caller expects no result.
	 * @return Whether the function executed successfully.
	*/
	HRESULT STDMETHODCALLTYPE Invoke(
		_In_  DISPPARAMS* pDispParams,
		_Out_ VARIANT*    pVarResult
	);

	/**
	 * @brief Dispatch ID associated to the dynamic method.
	*/
	DWORD m_dwDispatchId;

	/**
	 * @brief Name of the function to execute.
	*/
	BSTR m_bstrFunctionName;

	/**
	 * @brief Address of the function to execute.
	*/
	LPVOID m_lpFunction;
};

/**
 * @brief Dynamically execute a function. Function has been written in NASM x64.
 * @param lpTable The address of the table that contains the argument to pass to the function.
 * @param lpFunction The address of the function to execute.
 * @param lpResut The address of the RESULT union to return.
 * @param dwReturnFlag Whether scalar or non-scalar data is expected to be returned.
 * @return Whether the function executed successfully.
*/
extern "C" BOOL DynamicCall(
	_In_  PArgumentTable lpTable, 
	_In_  LPVOID         lpFunction, 
	_Out_ PRESULT        lpResut,
	_In_  DWORD          dwReturnFlag
);

#endif // !__DYNAMICMETHOD_HPP
