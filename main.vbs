'v1.1.0
option explicit
randomize
dim objShell, objFile, objHttp
set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
set objHttp = CreateObject("Msxml2.XMLHTTP.6.0")
dim directory

sub Main()
    dim config, url, notify, source, file, content, hasUpdate
    directory = objFile.GetParentFolderName(wscript.ScriptFullName)
    config = objFile.OpenTextFile(directory & "\config.txt").ReadAll
    url = KeyValueGet(config, "url")
    notify = KeyValueGet(config, "notify")
    source = KeyValueGet(config, "source")
    file = directory & "\temp.txt"
    hasUpdate = CheckForUpdates(source)
    content = Curl(url, 0, "")
    objFile.CreateTextFile(file, true).WriteLine(content)
    wscript.sleep(100)
    call objShell.Run("""" & file & """", 1, true)
    content = objFile.OpenTextFile(file).ReadAll
    call Curl(url, 1, content)

    if notify = 1 then
        call msgbox("Update successfull", 64, "notepad.exe")
    end if
end sub

sub Breakpoint(input)
    wscript.echo(input)
    wscript.quit
end sub

sub Debug()
    objFile.CreateTextFile("C:\Users\138670\Desktop\online notes.txt")
    wscript.quit
end sub

function Curl(url, method, data)
    if method = 0 then
        call objHttp.open("GET", url, false)
        call objHttp.setRequestHeader("Content-Type", "application/json")
        call objHttp.send()
	    Curl = objHttp.responseText
        exit function
    elseif method = 1 then
        call objHttp.open("PUT", url, false)
        call objHttp.setRequestHeader("Content-Type", "application/json")
        call objHttp.send(data)
        Curl = objHttp.responseText
        exit function
    else
        Curl = false
        exit function
    end if

    Curl = false
end function

function KeyValueGet(input, keyFind)
    dim pairs, i, element, key
    pairs = split(input, vbcrlf)

    for i = 0 to ubound(pairs)
        element = pairs(i)
        key = left(element, instr(element, ":") - 1)
        key = trim(key)
        
        if key = keyFind then
            KeyValueGet = mid(element, instr(element, ":") + 1)
            KeyValueGet = trim(KeyValueGet)
            exit function
        end if
    next
end function

function CheckForUpdates(url)
    dim sourceCode, currentCode, sourceVersion, currentVersion, x
    url = url & "?i=" & int(rnd * 1000000)
    currentCode = objFile.OpenTextFile(wscript.ScriptFullName).ReadAll
    currentCode = trim(currentCode)
    sourceCode = Curl(url, 0, "")
    sourceCode = trim(sourceCode)
    currentVersion = split(currentCode, vbcrlf)(0)
    currentVersion = mid(currentVersion, 2)
    sourceVersion = split(sourceCode, vbcrlf)(0)
    sourceVersion = mid(sourceVersion, 2)

    if currentVersion = sourceVersion then
    else
        x = msgbox("There's a new version available." & vbcrlf & "Would you like to update?", 64+4)

        if x = 6 then
            objFile.CreateTextFile(wscript.ScriptFullName, true).WriteLine(sourceCode)
            wscript.sleep(100)
            objShell.Run("""" & wscript.ScriptFullName & """")
            wscript.quit
        end if
    end if
end function

'Debug()
Main()
