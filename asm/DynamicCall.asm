; @file        DynamicCall.asm
; @date        01-10-2020
; @author      Paul Laîné (@am0nsec)
; @version     1.0
; @brief       Dynamically Execute Windows API.
; @details	
; @link        
; @copyright   This project has been released under the GNU Public License v3 license.
BITS 64
DEFAULT REL

%include "DynamicCall.inc"
global DynamicCall

;--------------------------------------------------------------------------------------------------
; DynamicCall procedure
;--------------------------------------------------------------------------------------------------
section .text
DynamicCall:
    %push mycontext             ; save the current context 
    %stacksize flat64            ; tell NASM to use bp 
    %assign %$localsize 0
    %local lpFunction:qword, lpTableArguments:qword, dwArguments:dword, lpResult:qword, dwReturnFlag:dword

    enter %$localsize, 0x00 ;
    push rbp                ; 
    push rbx                ;
    push rsi                ;
    push rdi                ;
    push r15                ;
    push r14                ;
    push r13                ;
    push r12                ;
    push r11                ;
    sub rsp, 200h           ; Reserve 512 bytes on the stack

;--------------------------------------------------------------------------------------------------
; Fill local stack variables
;--------------------------------------------------------------------------------------------------
    mov [lpFunction], rdx                              ; Second parameter
    mov [lpResult], r8                                 ; Address to the RESULT.
    mov [dwReturnFlag], r9d                            ; Whether the data to return is scalar

    mov r12, rcx                                       ;
    mov r11d, dword [r12 + ArgumentTable.dwArguments]  ; Get number of parameters to parse
    mov [dwArguments], r11d                            ;

    mov r11, [r12 + ArgumentTable.lpArguments]         ; Get the address of the list of parameters
    mov [lpTableArguments], r11                        ;

;--------------------------------------------------------------------------------------------------
; Parse each argument one by one
;--------------------------------------------------------------------------------------------------
.parse_arguments:
    mov eax, dword [dwArguments]      ; Total number of arguments 
    dec rax                           ; Index start to 0, hence -1 
    imul rax, rax, 10h                ; Calculate argument index

    add rax, qword [lpTableArguments] ; EAX = PArgument
    mov ebx, [rax + Argument.dwSize]  ; 

    cmp dword [dwArguments], 4h       ; Check if shadow stack or actual stack
    jle .shadow_stack                 ;

;--------------------------------------------------------------------------------------------------
; Allocate data to stack
;--------------------------------------------------------------------------------------------------
    mov ebx, dword [dwArguments]                    ; EBX = number of argumetns
    sub rbx, 5h                                     ; EBX - shadow stack parameters - 1
    imul rbx, rbx, 8h                               ; EBX = 8 * EBX

    add rbx, 20h                                    ; EBX = EBX + 20h
    add rbx, rsp                                    ; EBX = position in the stack

    and dword [rax + Argument.dwFlag], ARGUMENT_FLT ; Is non-scalar value
    jnz .stack_xmmx
    mov r10, [rax + Argument.value]                 ; Value as scalar data
    mov qword [rbx], r10                            ; Push value to the stack
    jmp .next                                       ;
.stack_xmmx:
    movsd xmm5, [rax + Argument.value]              ; Value as non-scalar data
    movsd [rbx], xmm5                               ; Push value to the stack
    jmp .next                                       ;

;--------------------------------------------------------------------------------------------------
; Allocate data to shadow stack
;--------------------------------------------------------------------------------------------------
.shadow_stack:
    mov r10d, dword [dwArguments]                   ;
    cmp r10d, 4h                                    ; 1 parameter
    je .param4                                      ; 
    cmp r10d, 3h                                    ; 2 parameter
    je .param3                                      ;
    cmp r10d, 2h                                    ; 3 parameter
    je .param2                                      ;
    cmp r10d, 1h                                    ; 4 parameter
    je .param1                                      ;
    jmp .failure                                    ; Something went wrong at this point
.param1:
    and dword [rax + Argument.dwFlag], ARGUMENT_FLT
    jnz .param1_xmmx
    mov rcx, [rax + Argument.value]
    jmp .next
.param1_xmmx:
    movsd xmm0, [rax + Argument.value]              ; floating point value 
    jmp .next                                       ;
.param2:
    and dword [rax + Argument.dwFlag], ARGUMENT_FLT ;
    jnz .param2_xmmx                                ;
    mov rdx, [rax + Argument.value]                 ; floating point value 
    jmp .next
.param2_xmmx:
    movsd xmm1, [rax + Argument.value]              ; floating point value 
    jmp .next                                       ;
.param3:
    and dword [rax + Argument.dwFlag], ARGUMENT_FLT ;
    jnz .param3_xmmx                                ;
    mov r8, [rax + Argument.value]                  ; floating point value 
    jmp .next
.param3_xmmx:
    movsd xmm2, [rax + Argument.value]              ; floating point value 
    jmp .next                                       ;
.param4:
    and dword [rax + Argument.dwFlag], ARGUMENT_FLT ;
    jnz .param4_xmmx                                ;
    mov r9, [rax + Argument.value]                  ;
    jmp .next
.param4_xmmx:
    movsd xmm3, [rax + Argument.value]              ; floating point value 
    jmp .next                                       ;
.next:
    dec dword [dwArguments]                         ; Next argument
    jnz .parse_arguments                            ; 

;--------------------------------------------------------------------------------------------------
; Call function
;--------------------------------------------------------------------------------------------------
    mov rax, qword [lpFunction] ; Address of the function
    call rax                    ;

;--------------------------------------------------------------------------------------------------
; Get return value
;--------------------------------------------------------------------------------------------------
    mov rbx, qword [lpResult]            ; Address of the RESULT union
    and dword [dwReturnFlag], RETURN_FLT ; Whether non-scalar return value
    jnz .result_xmmx                     ;
    mov qword [rbx], rax                 ; Move scalar data
    jmp .exit                            ;
.result_xmmx:
    movsd [rbx], xmm0                    ; Move non-scalar data
    jmp .exit                            ;

;--------------------------------------------------------------------------------------------------
; Failure
;--------------------------------------------------------------------------------------------------
.failure:
    xor eax, eax  ; FALSE
    jmp .exit     ;

;--------------------------------------------------------------------------------------------------
; Exit procedure
;--------------------------------------------------------------------------------------------------
.exit: 
    add rsp, 200h ;
    pop r11       ;
    pop r12       ;
    pop r13       ;
    pop r14       ;
    pop r15       ;
    pop rdi       ;
    pop rsi       ;
    pop rbx       ;
    pop rbp       ;

    leave 
    ret           ;
    %pop