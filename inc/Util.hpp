/**
* @file			Util.hpp
* @date			01-10-2020
* @author		Paul Laine (@am0nsec)
* @version		1.0
* @brief		Utility methods definition.
* @details
* @link         https://github.com/am0nsec/DynamicWrapperEx
* @copyright	This project has been released under the GNU Public License v3 license.
*/
#pragma once
#include <windows.h>
#include <OAIdl.h>

#ifndef __UTIL_HPP
#define __UTIL_HPP

// Add macro for Variants
#ifdef __clang__
// TODO: need to investigate the use of VARIANT (aka tagVARIANT) with clang-cl.
#else
#define __v_pbstr(x) x.pbstrVal
#define __v_bstr(x) x.bstrVal
#endif // !__clang__

/**
 * @brief Utility class.
*/
class Util {
public:

	/**
	 * @brief Write a byte at a given location.
	 * @param pDispParams Pointer to a DISPPARAMS structure containing an array of arguments provided by the client.
	 * @param pVarResult Pointer to the location where the result is to be stored, or NULL if the caller expects no result.
	 * @return Whether the function executed successfully.
	*/
	static HRESULT STDMETHODCALLTYPE WriteByte(
		_In_  DISPPARAMS* pDispParams,
		_Out_ VARIANT*    pVarResult
	);
};

#endif // !__UTIL_HPP
