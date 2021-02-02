/**
* @file			dllmain.cpp
* @date			01-10-2020
* @author		Paul Laine (@am0nsec)
* @version		1.0
* @brief		Windows module entry point.
* @details
* @link			https://github.com/am0nsec/DynamicWrapperEx
* @copyright	This project has been released under the GNU Public License v3 license.
*/
#include <windows.h>
#include <memory>

#include "CDynamicWrapperEx.hpp"

/**
 * @brief Dynamic-Link Library entry point.
 * @param hModule The address of the module in memory.
 * @param dwReasonToCall Notification.
 * @param lpReserved Reserved.
 * @return Whether the function executed successfully.
*/
BOOL STDMETHODCALLTYPE DllMain(
	_In_ HMODULE hModule,
	_In_ DWORD dwReasonForCall,
	_In_ LPVOID lpReserved
) {
	UNREFERENCED_PARAMETER(hModule);
	UNREFERENCED_PARAMETER(lpReserved);

	// Disables the DLL_THREAD_ATTACH and DLL_THREAD_DETACH notifications
	if (dwReasonForCall == DLL_PROCESS_ATTACH)
		::DisableThreadLibraryCalls(hModule);

	return TRUE;
}

/**
 * @brief  Check if the module can be unloaded.
 * @return Whether the module can be unloaded.
*/
HRESULT STDMETHODCALLTYPE DllCanUnloadNow(VOID) {
	return S_OK;
}

/**
 * @brief  Register the COM server into the Windows Registry.
 * @return Whether the COM server has been successfully registered.
*/
HRESULT STDMETHODCALLTYPE DllRegisterServer(VOID) {
	return S_OK;
}

/**
 * @brief  Unregister the COM server into the Windows Registry.
 * @return Whether the COM server has been successfully unregistered.
*/
HRESULT STDMETHODCALLTYPE DllUnregisterServer(VOID) {
	return S_OK;
}

/**
 * @brief  Get a COM class factory to create a COM object.
 * @param  rclsid GUID of the COM class factory.
 * @param  riid GUID of the COM server interface.
 * @param  ppv Address of the COM object.
 * @return Whether the COM object has been successfully created.
*/
HRESULT STDMETHODCALLTYPE DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
	if (!IsEqualGUID(rclsid, CLSID_CDynamicWrapperEx))
		return CLASS_E_CLASSNOTAVAILABLE;
	
	CDynamicWrapperEx* pClassFactory = new CDynamicWrapperEx;
	if (pClassFactory == nullptr)
		return E_OUTOFMEMORY;

	HRESULT hr = pClassFactory->QueryInterface(riid, ppv);
	pClassFactory->Release();
	return hr;
}