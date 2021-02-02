/**
* @file         AutomationFactory.hpp
* @date         01-10-2020
* @author       Paul Laine (@am0nsec)
* @version      1.0
* @brief        COM Automation Factory declaration.
* @details
* @link         https://github.com/am0nsec/DynamicWrapperEx
* @copyright    This project has been released under the GNU Public License v3 license.
*/
#pragma once
#include <windows.h>
#include <memory>
#include <vector>

#include "DynamicMethod.hpp"

#ifndef __AUTOMATIONFACTORY_HPP
#define __AUTOMATIONFACTORY_HPP

class AutomationFactory {
public:
	/**
	 * @brief Constructor.
	*/
	AutomationFactory();

	/**
	 * @brief Destructor.
	*/
	~AutomationFactory();

	/**
	 * @brief Register a new dynamic method.
	 * @param pDispParams List of parameters supplied by the client.
	 * @param pVarResult Return value expected by the client, if not NULL.
	 * @return Whether the function executed successfully.
	*/
	HRESULT STDMETHODCALLTYPE Register(
		_In_  DISPPARAMS* pDispParams,
		_Out_ VARIANT*    pVarResult
	);

	/**
	 * @brief Number of dynamic method that have been registered.
	*/
	DWORD m_dwDynamicMethods{ 0 };

	/**
	 * @brief Number of internal methods.
	*/
	DWORD m_dwInternalMethods{ 0 };

	/**
	 * @brief Table of dynamic methods
	*/
	std::vector<std::unique_ptr<DynamicMethod>> m_aDynamicMethods{};

	/**
	 * @brief List of Dispatch function.
	*/
	std::vector<DispatchTableEntry> m_aDispatchTable{};
private:
	/**
	 * @brief Get the address of a function from a module.
	 * @param pbstrModuleName The name of the module (e.g. user32.dll).
	 * @param pbstrFunctionName The name of the function (e.g. MessageBoxW).
	 * @param ppFunction The address of pointer variable that receives the address of the function.
	 * @return Whether the address of the function has been found.
	*/
	HRESULT STDMETHODCALLTYPE GetFunctionFromModule(
		_In_  BSTR*   pbstrModuleName,
		_In_  BSTR*   pbstrFunctionName,
		_Out_ LPVOID* ppFunction
	);
};

#endif // !__AUTOMATIONFACTORY_HPP
