#NoEnv  ;Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ;Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ;Ensures a consistent starting directory.
#SingleInstance Force ;Makes it so only one script instance is running at any time and skips prompt.


; This script is only the AHK part (1/2) of the full script, the other half is at autoImageSquarer.py.

PrintScreen & s::
    ; Prompt text box for user to enter category of squaring images. 
    InputBox, strCategory, Auto Image Squarer, Select Category:  [p/l/manual], , 210, 130, , , , ,
    RunAutoImageSquarer(strCategory)


RunAutoImageSquarer(argument)
{
    q := """" ;Chr(34) ;Double Quote char => "
    If ErrorLevel = 1  
    {
        Return  ; User presses "Cancel" button.
    }
    Else If (!argument || argument = "manual")
    {
        pathManual := ChooseFolder( [0, "Open"], A_Desktop, [], 0x10000000)
        If (pathManual = 0)  
        {
            Return  ; Do nothing sif no directory is selected
        }
        emptyArgument := "manual"
        commands := "python " q . "C:\Users\Admin\Documents\VS Code\AHK\Startup\AutoImageSquarer - AHK\AutoImageSquarer - Python.py" . q " " q . emptyArgument . q " " q . pathManual . q . "`n"
    }
    Else
    {
        commands := "python " q . "C:\Users\Admin\Documents\VS Code\AHK\Startup\AutoImageSquarer - AHK\AutoImageSquarer - Python.py" . q " " q . argument . q . "`n"
    }
    Run, cmd /c %commands%,, Hide
}

; ==================================================================================

; Author: Flipeador 20180803
; Ori Code: https://www.autohotkey.com/boards/viewtopic.php?f=76&t=53136&p=231879

; ==================================================================================

; Displays a standard dialog that allows the user to select folder(s).
; Parameters:
;     Owner / Title:
;         The identifier of the window that owns this dialog. This value can be zero.
;         An Array with the identifier of the owner window and the title. If the title is an empty string, it is set to the default.
;     StartingFolder:
;         The path to the directory selected by default. If the directory does not exist, it searches in higher directories.
;     CustomPlaces:
;         Specify an Array with the custom directories that will be displayed in the left pane. Missing directories will be omitted.
;         To specify the location in the list, specify an Array with the directory and its location (0 = Lower, 1 = Upper).
;     Options:
;         Determines the behavior of the dialog. This parameter must be one or more of the following values.
;         0x00000200 (FOS_ALLOWMULTISELECT) = Enables the user to select multiple items in the open dialog.
;         0x00040000 (FOS_HIDEPINNEDPLACES) = Hide items shown by default in the view's navigation pane.
;         0x02000000  (FOS_DONTADDTORECENT) = Do not add the item being opened or saved to the recent documents list (SHAddToRecentDocs).
;         0x10000000  (FOS_FORCESHOWHIDDEN) = Include hidden and system items.
;         You can check all available values ​​at https://msdn.microsoft.com/en-us/library/windows/desktop/dn457282(v=vs.85).aspx.
; Return:
;     Returns zero if the user canceled the dialog, otherwise returns the path of the selected directory. The directory never ends with "\".

; ========================================================================

ChooseFolder(Owner, StartingFolder := "", CustomPlaces := "", Options := 0)
{
    ; IFileOpenDialog interface
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb775834(v=vs.85).aspx
    local IFileOpenDialog := ComObjCreate("{DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7}", "{D57C7288-D4AD-4768-BE02-9D969532D960}")
        ,           Title := IsObject(Owner) ? Owner[2] . "" : ""
        ,           Flags := 0x20 | Options    ; FILEOPENDIALOGOPTIONS enumeration (https://msdn.microsoft.com/en-us/library/windows/desktop/dn457282(v=vs.85).aspx)
        ,      IShellItem := PIDL := 0         ; PIDL recibe la dirección de memoria a la estructura ITEMIDLIST que debe ser liberada con la función CoTaskMemFree
        ,             Obj := {}, foo := bar := ""
    Owner := IsObject(Owner) ? Owner[1] : (WinExist("ahk_id" . Owner) ? Owner : 0)
    CustomPlaces := IsObject(CustomPlaces) || CustomPlaces == "" ? CustomPlaces : [CustomPlaces]


    while (InStr(StartingFolder, "\") && !DirExist(StartingFolder))
        StartingFolder := SubStr(StartingFolder, 1, InStr(StartingFolder, "\",, -1) - 1)
    if ( DirExist(StartingFolder) )
    {
        StrPutVar(StartingFolder, StartingFolderW, "UTF-16")
        DllCall("Shell32.dll\SHParseDisplayName", "UPtr", &StartingFolderW, "Ptr", 0, "UPtrP", PIDL, "UInt", 0, "UInt", 0)
        DllCall("Shell32.dll\SHCreateShellItem", "Ptr", 0, "Ptr", 0, "UPtr", PIDL, "UPtrP", IShellItem)
        ObjRawSet(Obj, IShellItem, PIDL)
        ; IFileDialog::SetFolder method
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761828(v=vs.85).aspx
        DllCall(NumGet(NumGet(IFileOpenDialog+0)+12*A_PtrSize), "Ptr", IFileOpenDialog, "UPtr", IShellItem)
    }


    if ( IsObject(CustomPlaces) )
    {
        local Directory := ""
        For foo, Directory in CustomPlaces    ; foo = index
        {
            foo := IsObject(Directory) ? Directory[2] : 0    ; FDAP enumeration (https://msdn.microsoft.com/en-us/library/windows/desktop/bb762502(v=vs.85).aspx)
            if ( DirExist(Directory := IsObject(Directory) ? Directory[1] : Directory) )
            {
                StrPutVar(Directory, DirectoryW, "UTF-16")
                DllCall("Shell32.dll\SHParseDisplayName", "UPtr", &DirectoryW, "Ptr", 0, "UPtrP", PIDL, "UInt", 0, "UInt", 0)
                DllCall("Shell32.dll\SHCreateShellItem", "Ptr", 0, "Ptr", 0, "UPtr", PIDL, "UPtrP", IShellItem)
                ObjRawSet(Obj, IShellItem, PIDL)
                ; IFileDialog::AddPlace method
                ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb775946(v=vs.85).aspx
                DllCall(NumGet(NumGet(IFileOpenDialog+0)+21*A_PtrSize), "UPtr", IFileOpenDialog, "UPtr", IShellItem, "UInt", foo)
            }
        }
    }

    ; IFileDialog::SetTitle method
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761834(v=vs.85).aspx
    StrPutVar(Title, TitleW, "UTF-16")
    DllCall(NumGet(NumGet(IFileOpenDialog+0)+17*A_PtrSize), "UPtr", IFileOpenDialog, "UPtr", Title == "" ? 0 : &TitleW)

    ; IFileDialog::SetOptions method
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761832(v=vs.85).aspx
    DllCall(NumGet(NumGet(IFileOpenDialog+0)+9*A_PtrSize), "UPtr", IFileOpenDialog, "UInt", Flags)


    ; IModalWindow::Show method
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761688(v=vs.85).aspx
    local Result := []
    if ( !DllCall(NumGet(NumGet(IFileOpenDialog+0)+3*A_PtrSize), "UPtr", IFileOpenDialog, "Ptr", Owner, "UInt") )
    {
        ; IFileOpenDialog::GetResults method
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb775831(v=vs.85).aspx
        local IShellItemArray := 0    ; IShellItemArray interface (https://msdn.microsoft.com/en-us/library/windows/desktop/bb761106(v=vs.85).aspx)
        DllCall(NumGet(NumGet(IFileOpenDialog+0)+27*A_PtrSize), "UPtr", IFileOpenDialog, "UPtrP", IShellItemArray)

        ; IShellItemArray::GetCount method
        ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761098(v=vs.85).aspx
        local Count := 0    ; pdwNumItems
        DllCall(NumGet(NumGet(IShellItemArray+0)+7*A_PtrSize), "UPtr", IShellItemArray, "UIntP", Count)

        local Buffer := ""
        VarSetCapacity(Buffer, 32767 * 2)
        loop % Count
        {
            ; IShellItemArray::GetItemAt method
            ; https://msdn.microsoft.com/en-us/library/windows/desktop/bb761100(v=vs.85).aspx
            DllCall(NumGet(NumGet(IShellItemArray+0)+8*A_PtrSize), "UPtr", IShellItemArray, "UInt", A_Index-1, "UPtrP", IShellItem)

            ; IShellItem::GetDisplayName method
            ; https://docs.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishellitem-getdisplayname
            DllCall(NumGet(NumGet(IShellItem+0)+5*A_PtrSize), "Ptr", IShellItem, "Int", 0x80028000, "PtrP", ptr:=0)
            ObjRawSet(Obj, IShellItem, ptr), ObjPush(Result, RTrim(StrGet(ptr,"UTF-16"), "\"))

            if (Result[A_Index] ~= "^::")  ; handle "::{00000000-0000-0000-0000-000000000000}\Documents.library-ms" (library)
            {
            	; SHLoadLibraryFromParsingName
            	; https://docs.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-shloadlibraryfromparsingname
                VarSetCapacity(IID_IShellItem, 16)
                DllCall("Ole32\CLSIDFromString", "Str", "{43826d1e-e718-42ee-bc55-a1e261c37bfe}", "Ptr", &IID_IShellItem)
                DllCall("Shell32\SHCreateItemFromParsingName", "WStr", Result[A_Index], "Ptr", 0, "Ptr", &IID_IShellItem, "PtrP", IShellItem:=0)

                IShellLibrary := ComObjCreate("{d9b3211d-e57f-4426-aaef-30a806add397}", "{11A66EFA-382E-451A-9234-1E0E12EF3085}")
                ; IShellLibrary::LoadLibraryFromItem
                ; https://docs.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-ishelllibrary
                DllCall(NumGet(NumGet(IShellLibrary+0)+3*A_PtrSize), "UPtr", IShellLibrary, "Ptr", IShellItem, "Int", 0)
                IShellLibrary2 := ComObjQuery(IShellLibrary, "{11A66EFA-382E-451A-9234-1E0E12EF3085}")
                ObjRelease(IShellLibrary)
                ObjRelease(IShellItem)

                VarSetCapacity(IID_IShellItemArray, 16)
                DllCall("Ole32\CLSIDFromString", "Str", "{b63ea76d-1f85-456f-a19c-48159efa858b}", "Ptr", &IID_IShellItemArray)
                ; IShellLibrary::GetFolders method
                ; https://docs.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-ishelllibrary-getfolders
                DllCall(NumGet(NumGet(IShellLibrary2+0)+7*A_PtrSize), "UPtr", IShellLibrary2, "Int", 1, "Ptr", &IID_IShellItemArray, "PtrP", IShellItemArray:=0)

                DllCall(NumGet(NumGet(IShellItemArray+0)+8*A_PtrSize), "UPtr", IShellItemArray, "Int", 0, "PtrP", IShellItem0:=0)
                DllCall(NumGet(NumGet(IShellItem0+0)+5*A_PtrSize), "Ptr", IShellItem0, "Int", 0x80028000, "PtrP", ptr:=0)
                Result[A_Index] := StrGet(ptr, "UTF-16")
                DllCall("Ole32\CoTaskMemFree", "Ptr", ptr)
                ObjRelease(IShellItem0)
                ObjRelease(IShellItemArray)
                ObjRelease(IShellLibrary2)
            }
        }

        ObjRelease(IShellItemArray)
    }


    for foo, bar in Obj    ; foo = IShellItem interface (ptr)  |  bar = PIDL structure (ptr)
        ObjRelease(foo), DllCall("Ole32.dll\CoTaskMemFree", "Ptr", bar)
    ObjRelease(IFileOpenDialog)

    return ObjLength(Result) ? ( Options & 0x200 ? Result : Result[1] ) : FALSE
}

DirExist(DirName)
{
    loop Files, % DirName, D
        return A_LoopFileAttrib
}

StrPutVar(string, ByRef var, encoding)
{
    ; Ensure capacity.
    VarSetCapacity( var, StrPut(string, encoding)
        ; StrPut returns char count, but VarSetCapacity needs bytes.
        * ((encoding="utf-16"||encoding="cp1200") ? 2 : 1) )
    ; Copy or convert the string.
    return StrPut(string, &var, encoding)
} ; https://www.autohotkey.com/docs/commands/StrPut.htm#Examples

; ========================================================================