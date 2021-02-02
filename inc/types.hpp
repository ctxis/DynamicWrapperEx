/**
* @file			Types.hpp
* @date			01-10-2020
* @author		Paul Laine (@am0nsec)
* @version		1.0
* @brief		Custom types.
* @details
* @link         https://github.com/am0nsec/DynamicWrapperEx
* @copyright	This project has been released under the GNU Public License v3 license.
*/
#include <windows.h>
#include <functional>

#ifndef __TYPES_HPP
#define __TYPES_HPP

typedef struct _DispatchTableEntry {
	DISPID  lDispId;
	LPWSTR  wszName;
} DispatchTableEntry, *PDispatchTableEntry;

#pragma pack(push)
#pragma pack(1)
/**
 * @brief Argument to pass to the function. Pack(1) to be used in NASM x64.
*/
typedef struct _Argument {
	DWORD dwFlag;
	DWORD dwSize;
	union {
		DWORD64 qwValue;
		LPVOID  lpValue;
		DOUBLE  rlValue;
	};
} Argument, *PArgument;

/**
 * @brief Table of argument to pass to the function. Pack(1) to be used in NASM x64.
*/
typedef struct _ArgumentTable {
	DWORD     dwArguments;
	Argument* lpArguments;
} ArgumentTable, * PArgumentTable;

/**
 * @brief Contains standard/floating point return value from the function. Pack(1) to be used in NASM x64.
*/
typedef union RESULT {
	INT     inValue;
	LONG    lgValue;
	LPVOID  lpValue;
	FLOAT   flValue;
	DOUBLE  dbValue;
	LONGLONG int64;
} RESULT, * PRESULT;
#pragma pack (pop)

#endif // !__TYPES_HPP
