/**
* @file         AutomationFactory.cpp
* @date         01-10-2020
* @author       Paul Laine (@am0nsec)
* @version      1.0
* @brief        COM Automation Factory definition.
* @details
* @link         https://github.com/am0nsec/DynamicWrapperEx
* @copyright    This project has been released under the GNU Public License v3 license.
*/
#include <windows.h>
#include <comutil.h>
#include <string>
#include <memory>
#include <vector>

#include "AutomationFactory.hpp"
#include "DynamicMethod.hpp"
#include "Util.hpp"

/**
 * @brief Constructor
*/
AutomationFactory::AutomationFactory() { };

/**
 * @brief Destructor
*/
AutomationFactory::~AutomationFactory() {
	this->m_aDynamicMethods.clear();
	this->m_aDispatchTable.clear();
}

/**
 * @brief Register a new dynamic method.
 * @param pDispParams List of parameters supplied by the client.
 * @param pVarResult Return value expected by the client, if not NULL.
 * @return Whether the function executed successfully.
*/
HRESULT STDMETHODCALLTYPE AutomationFactory::Register(
	_In_  DISPPARAMS* pDispParams,
	_Out_ VARIANT*    pVarResult
) {
	// Check number of arguments
	if (pDispParams->cArgs != 2)
		return E_FAIL;

	// Get parameters
	BSTR* pbstrFunctionName = &pDispParams->rgvarg[0].bstrVal;
	BSTR* pbstrModuleName = &pDispParams->rgvarg[1].bstrVal;

	// Get function address
	LPVOID lpFunction = NULL;
	if (this->GetFunctionFromModule(pbstrModuleName, pbstrFunctionName, &lpFunction) != S_OK || lpFunction == NULL)
		return E_FAIL;

	// Create new dynamic method
	std::unique_ptr<DynamicMethod> dm = std::make_unique<DynamicMethod>(this->m_dwDynamicMethods, *pbstrFunctionName, lpFunction);
	this->m_aDynamicMethods.push_back(std::move(dm));
	this->m_aDispatchTable.push_back({ static_cast<DISPID>(this->m_dwDynamicMethods + this->m_dwInternalMethods), *pbstrFunctionName });
	this->m_dwDynamicMethods++;

	if (pVarResult) {
		V_VT(pVarResult) = VT_BOOL;
		V_BOOL(pVarResult) = TRUE;
	}
	return S_OK;
}

/**
 * @brief Get the address of a function from a module.
 * @param pbstrModuleName The name of the module (e.g. user32.dll).
 * @param pbstrFunctionName The name of the function (e.g. MessageBoxW).
 * @param ppFunction The address of pointer variable that receives the address of the function.
 * @return Whether the address of the function has been found.
*/
HRESULT STDMETHODCALLTYPE AutomationFactory::GetFunctionFromModule(
	_In_  BSTR*   pbstrModuleName,
	_In_  BSTR*   pbstrFunctionName,
	_Out_ LPVOID* ppFunction
) {
	*ppFunction = NULL;

	// Check if values are empty
	if (::SysStringLen(*pbstrModuleName) == 0 || ::SysStringLen(*pbstrFunctionName) == 0)
		return E_FAIL;

	// Check if the module has been already loaded
	HMODULE hModule = NULL;
	hModule = ::LoadLibraryW(*pbstrModuleName);
	if (hModule == NULL)
		return E_FAIL;

	// Convert string name into ASCII
	std::string szFunctionName = static_cast<std::string>(_bstr_t(*pbstrFunctionName));
	FARPROC lpProcAddress = ::GetProcAddress(hModule, szFunctionName.c_str());
	if (lpProcAddress == nullptr)
		return E_FAIL;

	*ppFunction = reinterpret_cast<LPVOID>(lpProcAddress);
	return S_OK;
}