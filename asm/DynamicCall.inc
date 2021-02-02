; @file        DynamicCall.inc
; @date        01-10-2020
; @author      Paul Laîné (@am0nsec)
; @version     1.0
; @brief       Dynamically Execute Windows API.
; @details	
; @link        
; @copyright   This project has been released under the GNU Public License v3 license.

ARGUMENT_STD equ 0
ARGUMENT_FLT equ 2
RETURN_STD   equ 0
RETURN_FLT   equ 2

struc Argument
    .dwFlag: resd 1
    .dwSize: resd 1
    .value:  resq 1
endstruc

struc ArgumentTable
    .dwArguments: resd 1
    .lpArguments: resq 1
endstruc