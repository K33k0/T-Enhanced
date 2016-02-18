#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;run, update/git-pull.exe "../../"
shell := ComObjCreate("WScript.Shell")
shell.CurrentDirectory := A_ScriptDir . "\update"
exec := shell.Exec(ComSpec " /Q /k echo off")
commands := "
(
dir
git-pull.exe 
exit
)"
exec.StdIn.WriteLine(commands)
MsgBox % exec.StdOut.ReadAll()