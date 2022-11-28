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
#include <windows.h>
#include "Util.hpp"

/**
 * @brief Write a byte at a given location.
 * @param pDispParams Pointer to a DISPPARAMS structure containing an array of arguments provided by the client.
 * @param pVarResult Pointer to the location where the result is to be stored, or NULL if the caller expects no result.
 * @return Whether the function executed successfully.
*/
HRESULT STDMETHODCALLTYPE Util::WriteByte(
    _In_  DISPPARAMS* pDispParams,
    _Out_ VARIANT*    pVarResult
) {
    // Check number of arguments
    if (pDispParams->cArgs != 3) {
        if (pVarResult != NULL) {
            V_VT(pVarResult) = VT_BOOL;
            V_BOOL(pVarResult) = FALSE;
        }
        return E_FAIL;
    }

    // Get parameters

    BYTE bValue = V_UI1(&pDispParams->rgvarg[2]);
    PBYTE lpAddress = (PBYTE)V_UI8(&pDispParams->rgvarg[1]);
    DWORD dwOffset = V_I4(&pDispParams->rgvarg[0]);

    // Write byte
    *(lpAddress + dwOffset) = bValue;

    if (pVarResult != NULL) {
        V_VT(pVarResult) = VT_BOOL;
        V_BOOL(pVarResult) = TRUE;
    }

    return S_OK;
}
