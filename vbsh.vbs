' @author: Grok
' @license: MIT
'
' A simple interactive VBScript shell.

version = "0.1"

' Add all scripts you want to import at startup
scripts = array()

' Set to True if you want to auto import files in scripts array
allways_import_scripts = False

' Default logging level is INFO
LogLevel = 4
LogPrefix = True


ParseInputArgs()
Boot()


private function Boot()
    LogDebug("Boot()")

    if allways_import_scripts = True then
        ImportInitScript()
        Main()
    else
        if ubound(scripts) >= 0 then
            ' Only prompt user input if there is any predefined scripts
            LogInfo("Do you want to import the predefined scripts? (In this order)" & VbCrLf & "")
            for each script in scripts
                LogInfo(" * " + script)
            next

            wscript.echo ""
            wscript.stdout.write("(y / n / q): ")

            line = trim(wscript.stdin.readline)

            if line = "y" then
                ImportInitScript()
                Main()
            elseif line = "n" then
                Main()
            elseif line = "q" then
                Print("Bai :]")
                wscript.quit(0)
            else
                LogError("Invalid input... Try again!")
            end if
        else
            Main()
        end if
    end if
end function


private function ImportInitScript()
    LogDebug("ImportInitScript()")

    for each script in scripts
        LogDebug("Import Initscript: " + script)

        set sh  = createobject("WScript.Shell")
        set fso = createobject("Scripting.FileSystemObject")

        path = sh.ExpandEnvironmentStrings(script)
        scriptExists = fso.FileExists(path)

        Set sh  = Nothing
        Set fso = Nothing

        if scriptExists then
            Import(path)
        end if
    next
end function


private function Main()
    LogDebug("Main()")

    do while True
        wscript.stdout.write(">>> ")

        line = trim(wscript.stdin.readline)
        do while right(line, 2) = " _" or line = "_"
            line = rtrim(left(line, len(line)-1)) & " " & trim(wscript.stdin.readline)
        loop

        if lcase(line) = "?exit" then
            exit do
        end if

        if line = "?" then
            PrintHelp()
        elseif StartsWith(line, "?import ") = True then
            file = Replace(line, "?import ", "")
            Import(file)
        elseif StartsWith(line, "?version") = True then
            Print(version)
        elseif StartsWith(line, "?reimport") = True then
            ImportInitScript()
        else
            on error resume next
            Err.clear
            Execute(line)
            if Err.Number <> 0 then
                fedcba = trim(Err.Description & " (0x" & hex(Err.Number) & ")")

                if Err.Number = 13 then
                    abcdef = "if VarType(" & line & ") = 0 then" & VbCrLf & _
                             "     wscript.echo " & quotes("Object not initilized") & VbCrLf & _
                             "else" & VbCrLf & _
                             "     wscript.echo CStr(" & line & ")" & VbCrLf & _
                             "end if" & VbCrLf
                    ExecuteCode(abcdef)
                else
                    wscript.echo "Compile-Error: " + Err.Description
                end if
            end if
            on error goto 0
        end if
    loop
end function


private function ExecuteCode(ByRef s)
    LogDebug("ExecuteCode()")
    LogDebug(s)
    Execute(s)
end function


private function quotes(string)
    LogDebug("quotes(" & string & ")")
    quotes = chr(34) + string + chr(34)
end function


private function PrintHelp()
    LogDebug("PrintHelp()")

    ' Print a Help message.
    Print("VBSchell " + version)
    Print("")
    Print("   ?           Prints this Help")
    Print("   ?exit       Exit the shell")
    Print("   ?import     Prompts to import a .vbs script")
    Print("   ?reimport   Reimports the predefined scripts")
    Print("   ?version    Prints the version")
end function


' Import the first occurrence of the given filename from the working directory
' or any directory in the %PATH%.
'
' @param  filename   Name of the file to import.
private function Import(ByVal filename)
    LogDebug("Import(" & filename & ")")

    set fso = createobject("Scripting.FileSystemObject")
    set sh = createobject("WScript.Shell")

    filename = trim(sh.ExpandEnvironmentStrings(filename))
    If Not (left(filename, 2) = "\\" or mid(filename, 2, 2) = ":\") then
        ' filename is not absolute
        if not fso.FileExists(fso.GetAbsolutePathName(filename)) then
            ' file doesn't exist in the working directory => iterate over the
            ' directories in the %PATH% and take the first occurrence
            ' if no occurrence is found => use filename as-is, which will result
            ' in an error when trying to open the file
            for each dir in split(sh.ExpandEnvironmentStrings("%PATH%"), ";")
                if fso.FileExists(fso.BuildPath(dir, filename)) then
                    filename = fso.BuildPath(dir, filename)
                    exit for
                end if
            next
        end if
        filename = fso.GetAbsolutePathName(filename)
    end if

    if fso.FileExists(filename) then
        set file = fso.OpenTextFile(filename, 1, False)
        code = file.ReadAll()
        file.Close()

        LogInfo("Importing file: " + filename)
        ExecuteGlobal(code)
    else
        LogError("File Not Found on disk...")
    end if

    set fso = Nothing
    set sh = Nothing
end function


private function StartsWith(string, what)
    LogDebug("StartsWith(" & string & ", " & what & ")")

    if InStr(Trim(string), what) = 1 then
        StartsWith = True
    else
        StartsWith = False
    end if

    LogDebug("return StartsWith() = " & StartsWith)
end function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Input argument parting

private function ParseInputArgs()
    LogDebug("ParseInputArgs()")

    ' Parse all input arguments and set variables
    for i = 0 to wscript.Arguments.count - 1
        a = wscript.arguments.Item(i)
        if a = "--help" then
            ' Using docopt cli specs "https://github.com/docopt/docopt"
            Print("Usage:")
            Print(" vbsh.wsf [-v ... ] [--help]")
            Print(" vbsh.wsf --version")
            Print(" ")
            Print("Options:")
            Print(" --help      Prints this help")
            Print(" -v ...      Sets the logging level of this run. [Default: -vvvv --> INFO logging] [Supports level 1-5]")
            Print(" --version   Prints the version of the application")
            Print("")
            wscript.quit(0)
        elseif a = "--version" then
            wscript.echo version
            wscript.quit(0)
        elseif a = "-v" then
            ' Critical log level
            LogLevel = 1
        elseif a = "-vv" then
            ' Error log level
            LogLevel = 2
        elseif a = "-vvv" then
            ' Warning log level
            LogLevel = 3
        elseif a = "-vvvv" then
            ' Info log level
            LogLevel = 4
        elseif a = "-vvvvv" then
            ' Debug log level
            LogLevel = 5
        else
            wscript.echo "ERROR: Unknown input argument: " + a
            wscript.quit(0)
        end if
    next
end function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Logging functions'

private function LogDebug(msg)
    if LogLevel >= 5 then
        if LogPrefix = True then
            print("DEBUG: " & cstr(msg))
        else
            print(cstr(msg))
        end if
    end if
end function


private function LogInfo(msg)
    if LogLevel >= 4 then
        if LogPrefix = True then
            print("INFO: " & cstr(msg))
        else
            print(cstr(msg))
        end if
    end if
end function


private function LogWarning(msg)
    if LogLevel >= 3 then
        if LogPrefix = True then
            print("WARNING: " & cstr(msg))
        else
            print(cstr(msg))
        end if
    end if
end function


private function LogError(msg)
    if LogLevel >= 2 then
        if LogPrefix = True then
            print("ERROR: " & cstr(msg))
        else
            print(cstr(msg))
        end if
    end if
end function


private function LogCritical(msg)
    if LogLevel >= 1 then
        if LogPrefix = True then
            print("CRITICAL: " & cstr(msg))
        else
            print(cstr(msg))
        end if
    end if
end function


private function print(msg)
    wscript.echo msg
end function


private function p(msg)
    wscript.echo msg
end function
