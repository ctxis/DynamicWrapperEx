/**
* @file         DynamicMethod.cpp
* @date         01-10-2020
* @author       Paul Laine (@am0nsec)
* @version      1.0
* @brief        Dynamic method definition.
* @details
* @link         https://github.com/am0nsec/DynamicWrapperEx
* @copyright    This project has been released under the GNU Public License v3 license.
*/
#include <windows.h>
#include <memory>
#include <vector>

#include "DynamicMethod.hpp"

/**
 * @brief Constructor.
 * @param dwDispatchId The dispatch ID that has been associated to this dynamic method.
 * @param bstrFunctionName The name of the function to execute.
 * @param lpFunction The address of the function to execute.
*/
DynamicMethod::DynamicMethod(
	_In_ DWORD  dwDispatchId,
	_In_ BSTR   bstrFunctionName,
	_In_ LPVOID lpFunction
) {
	this->m_dwDispatchId = dwDispatchId;
	this->m_bstrFunctionName = bstrFunctionName;
	this->m_lpFunction = lpFunction;
}

/**
 * @brief Destructor.
*/
DynamicMethod::~DynamicMethod() {
	if (this->m_bstrFunctionName)
		::SysFreeString(this->m_bstrFunctionName);
}

/**
 * @brief Dynamically execute the function associated to this dynamic method.
 * @param pDispParams List of parameters provided by the client.
 * @param pVarResult Pointer to the location where the result is to be stored, or NULL if the caller expects no result.
 * @return Whether the function executed successfully.
*/
HRESULT STDMETHODCALLTYPE DynamicMethod::Invoke(
	_In_  DISPPARAMS* pDispParams,
	_Out_ VARIANT*    pVarResult
) {
	// Continuous memory
	Argument* args = (Argument*)::CoTaskMemAlloc(sizeof(Argument) * pDispParams->cArgs);

	// Parse parameters
	for (WORD cx = 0; cx < pDispParams->cArgs; cx++) {

               	args[pDispParams->cArgs - cx - 1].dwFlag = ARGUMENT_STD;

		switch (pDispParams->rgvarg[cx].vt) {
			// non-scalar value
		case VT_R4:
		case VT_R8:
		case VT_DECIMAL:
			args[pDispParams->cArgs - cx - 1].dwFlag = ARGUMENT_FLT;
			args[pDispParams->cArgs - cx - 1].rlValue = pDispParams->rgvarg[cx].dblVal;
			break;

		case VT_I2:
		case VT_UI2:
			args[pDispParams->cArgs - cx - 1].qwValue = pDispParams->rgvarg[cx].iVal;
			break;

		case VT_INT:
		case VT_UINT:
		case VT_I4:
		case VT_UI4:
			args[pDispParams->cArgs - cx - 1].qwValue = pDispParams->rgvarg[cx].intVal;
			break;

		case VT_PTR:
		case VT_I8:
		case VT_UI8:
			args[pDispParams->cArgs - cx - 1].qwValue = pDispParams->rgvarg[cx].ullVal;
			break;

		case VT_NULL:
		case VT_VOID:
			args[pDispParams->cArgs - cx - 1].qwValue = 0;
			break;

		default:
			args[pDispParams->cArgs - cx - 1].lpValue = pDispParams->rgvarg[cx].pvRecord;
			break;
		}
	}


	// Execute function
	RESULT res{ 0 };
	ArgumentTable Table = { pDispParams->cArgs, args };
	DynamicCall(&Table, this->m_lpFunction, &res, RETURN_STD);

	// Return value 
	if (pVarResult) {
		pVarResult->ullVal = (ULONGLONG)res.lpValue;
		pVarResult->vt = VT_UI8;
	}
	return S_OK;
}
