; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
;                     COMPILER DIRECTIVES
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
.586                            ;Create 32 bit code
.model flat, stdcall            ;32 bit memory model
option casemap :none            ;Case sensitive
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
;                     INCLUDE FILES
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
    ;
    include \masm32\include\windows.inc
    include \masm32\INCLUDE\dialogs.inc
    ;include \masm32\include\masm32rt.inc
    ;include \masm32\include\masm32.inc
    ;include     \masm32\include\masm32.inc
    ;includelib  \masm32\lib\masm32.lib
    ;
    ;include     \masm32\include\advapi32.inc
    include \masm32\include\kernel32.inc
    include \masm32\include\user32.inc
    include \masm32\include\advapi32.inc
    include \masm32\include\comdlg32.inc
    include \masm32\include\Comctl32.inc    ;Invoke InitCommonControls
    ;include \masm32\include\gdi32.inc
    ;include \masm32\include\Shell32.inc
    ;include \masm32\include\ole32.inc
    ;include \masm32\include\version.inc
    ;
    include     \masm32\include\gdi32.inc
    include     \masm32\include\gdiplus.inc
    include     \masm32\include\msimg32.inc
    include     \masm32\include\ole32.inc
    include     \masm32\include\oleaut32.inc
    include     \masm32\include\Shell32.inc
    include     \masm32\include\version.inc
    include     \masm32\include\winmm.inc
    ;
    ;includelib  \masm32\lib\advapi32.lib
    includelib \masm32\lib\kernel32.lib
    includelib \masm32\lib\user32.lib
    includelib \masm32\lib\advapi32.lib
    includelib \masm32\lib\comdlg32.lib
    includelib \masm32\lib\Comctl32.lib     ;Invoke InitCommonControls
    ;includelib \masm32\lib\gdi32.lib
    ;includelib \masm32\lib\Shell32.lib
    ;includelib \masm32\lib\ole32.lib
    ;includelib \masm32\lib\version.lib
    ;
    includelib \masm32\lib\gdi32.lib
    includelib  \masm32\lib\gdiplus.lib
    includelib  \masm32\lib\msimg32.lib
    includelib  \masm32\lib\ole32.lib
    includelib  \masm32\lib\oleaut32.lib
    includelib  \masm32\lib\Shell32.lib
    includelib  \masm32\lib\version.lib
    includelib  \masm32\lib\winmm.lib
    ;
    include macros.inc
    ;
    ; -------------------------------------------
    ; include the image library and include file
    ; -------------------------------------------
    include image.inc
    includelib image.lib
    ;
    Include     \masm32\CarabeZ\AA_Controls\AAGrafics.asm     ;needs msimg32.inc,oleaut32.inc,ole32.inc,gdiplus.inc
    Include     \masm32\CarabeZ\AA_Controls\AAControls.asm    ;needs winmm.inc,comctl32.inc,kernel32.inc
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;GLOBAL FUNCTIONS
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
    RunCmdAtMainEntry           PROTO ;Result 0 to exit, else call main.
    MainDlgProc                 PROTO :DWORD,:DWORD,:DWORD,:DWORD
    ForceWindowToTop            PROTO :DWORD
    ;
    AboutDlg                    PROTO :DWORD,:DWORD
    AboutDlgProc                PROTO :DWORD,:DWORD,:DWORD,:DWORD
    ;
    ;
    ConvertStaticToHyperlink    PROTO :DWORD,:DWORD
    _HyperlinkParentProc        PROTO :DWORD,:DWORD,:DWORD,:DWORD
    _HyperlinkProc              PROTO :DWORD,:DWORD,:DWORD,:DWORD
    ;
    Paint_Proc                  PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD
    ;
    ;
    CheckCommandLineArgs        PROTO
    ProccessParameters          PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD    ;pszApp,pszPara,pszWrkDir,dTerm,dUnhide,dShow,dTimeOut
    ReadIniFile                 PROTO :DWORD;hWnd
    WriteIniFile                PROTO :DWORD;hWnd
    ValidateIniValue            PROTO :DWORD,:DWORD ;pSource,dSize
    FileDialog                  PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD    ;hWnd,pszTitle,pszFilename,dszFileNameSize,pszFilter/SecString,dSizeFilter,dDirectory
    cbBrowseWindow              PROTO :DWORD,:DWORD,:DWORD,:DWORD
    DirFromPath                 PROTO :DWORD,:DWORD,:DWORD ;Origen "c:\*.*",Target "c:\" from "c:\*.*",Leader "*.*" from "c:\*.*"
    BuildTreeDir                PROTO :DWORD
    CopyFiles                   PROTO :DWORD,:DWORD,:DWORD,:DWORD           ;pSource,pTarget,dSubfolders,dOverWrite
    SingleCopy                  PROTO :DWORD,:DWORD,:DWORD                   ;pSource,pTarget,dOverWrite
    DirExist                    PROTO :DWORD                                ;pDestFolder
    FileExist                   PROTO :DWORD                                ;pszFilePath
    WildCardCopy                PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD    ;Pattern,Target,dSubFolders,dOverWrite,pPattern(FALSE=Regular TRUE=Pattern in Subfolders),dIsParent
    ;
    ShellExecuteWait            PROTO :DWORD,:DWORD,:DWORD,:DWORD,:DWORD,:DWORD    ;pszApplicationPath,pszParameters,pszWorkingDir,dShowWindow,dTimeOut,dForceWait
    FindApplicationProcessID    PROTO :DWORD,:DWORD,:DWORD                  ;pszApplicationPath,dWhatToDoIfFound,pOutWinHandle
    FindTopWindowFromProcessID  PROTO :DWORD,:DWORD                 ;dProcessID,dWhatToDoIfFound
    EnumWindowsProc             PROTO :DWORD,:DWORD                         ;hwinparent,lparam=userdefinedparam
    GetProcessMainThreadID      PROTO :DWORD ; hProcessID
    AscciiToDword               PROTO :DWORD                                ;pszCstring
    DwordToAsccii               PROTO :DWORD,:DWORD                         ;dword,pszCstring
    GetCmdLine                  PROTO :DWORD,:DWORD                         ;dArgNumber,pCstring
    GetArgNumber                PROTO :DWORD,:DWORD,:DWORD,:DWORD
    MemCopy                     PROTO :DWORD,:DWORD,:DWORD
    FillBuffer                  PROTO :DWORD,:DWORD,:BYTE
    ;
    SolveMacros                 PROTO :DWORD,:DWORD                         ;lpSourceStr, lpTarget
    ReplaceMacros               PROTO :DWORD,:DWORD,:DWORD                  ;lpSourceStr, lpPattern, lpTarget
    PasteMacro                  PROTO :DWORD,:DWORD,:DWORD                  ;PopUpSelection,EditHandle,EditID
    SubWndProc                  PROTO :DWORD,:DWORD,:DWORD,:DWORD
    StrLen                      PROTO :DWORD
    InString                    PROTO :DWORD,:DWORD,:DWORD
    GetPathOnly                 PROTO :DWORD,:DWORD       ;SRC,DEST
    ;
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
.const
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;EQUATES
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
    IDC_GROUP_PRE                       equ 01
    IDC_PROMPT_PRE_SOURCE               equ 02
    IDC_EDIT_PRE_SOURCE                 equ 03
    IDC_BUTTON_PRE_SOURCE_FILEDLG       equ 04
    IDC_BUTTON_PRE_SOURCE_OVERWR        equ 05
    IDC_PROMPT_PRE_TARGET               equ 06
    IDC_EDIT_PRE_TARGET                 equ 07
    IDC_BUTTON_PRE_TARGET_FILEDLG       equ 08
    IDC_BUTTON_PRE_TARGET_OVERWR        equ 09
    IDC_PROMPT_PRE_EXECUTE              equ 10
    IDC_EDIT_PRE_EXECUTE                equ 11
    IDC_BUTTON_PRE_EXECUTE_FILEDLG      equ 12
    IDC_PROMPT_PRE_PARAMS               equ 13
    IDC_EDIT_PRE_PARAMS                 equ 14
    IDC_BUTTON_PRE_PARAMS_FILEDLG       equ 15
    IDC_PROMPT_PRE_WRKDIR               equ 16
    IDC_EDIT_PRE_WRKDIR                 equ 17
    IDC_BUTTON_PRE_WRKDIR_FILEDLG       equ 18
    IDC_PROMPT_PRE_OPTIONS              equ 19
    IDC_CHECK_PRE_UNHIDE                equ 20
    IDC_CHECK_PRE_SHOW                  equ 21
    IDC_CHECK_PRE_ASK_BEFORE_EXEC       equ 22
    IDC_CHECK_PRE_GUI_CANCEL_EXEC       equ 23
    IDC_CHECK_PRE_TERM                  equ 24
    IDC_CHECK_PRE_COPY_AFTER            equ 25
    IDC_CHECK_PRE_ASK_COPY_AFTER        equ 26
    IDC_CHECK_PRE_GUI_COPY_AFTER        equ 27
    IDC_PROMPT_PRE_TIMEOUT              equ 28
    IDC_EDIT_PRE_TIMEOUT                equ 29
    IDC_SPINBOX_PRE_TIMEOUT             equ 30
    IDC_CHECK_PRE_COPY_BEFORE           equ 31
    IDC_CHECK_PRE_ASK_COPY_BEFORE       equ 32
    IDC_CHECK_PRE_GUI_COPY_BEFORE       equ 33
    ;
    IDC_GROUP_MAIN                      equ 34
    IDC_PROMPT_MAIN_SOURCE              equ 35
    IDC_EDIT_MAIN_SOURCE                equ 36
    IDC_BUTTON_MAIN_SOURCE_FILEDLG      equ 37
    IDC_BUTTON_MAIN_SOURCE_OVERWR       equ 38
    IDC_PROMPT_MAIN_TARGET              equ 39
    IDC_EDIT_MAIN_TARGET                equ 40
    IDC_BUTTON_MAIN_TARGET_FILEDLG      equ 41
    IDC_BUTTON_MAIN_TARGET_OVERWR       equ 42
    IDC_PROMPT_MAIN_EXECUTE             equ 43
    IDC_EDIT_MAIN_EXECUTE               equ 44
    IDC_BUTTON_MAIN_EXECUTE_FILEDLG     equ 45
    IDC_PROMPT_MAIN_PARAMS              equ 46
    IDC_EDIT_MAIN_PARAMS                equ 47
    IDC_BUTTON_MAIN_PARAMS_FILEDLG      equ 48
    IDC_PROMPT_MAIN_WRKDIR              equ 49
    IDC_EDIT_MAIN_WRKDIR                equ 50
    IDC_BUTTON_MAIN_WRKDIR_FILEDLG      equ 51
    IDC_PROMPT_MAIN_OPTIONS             equ 52
    IDC_CHECK_MAIN_UNHIDE               equ 53
    IDC_CHECK_MAIN_SHOW                 equ 54
    IDC_CHECK_MAIN_ASK_BEFORE_EXEC      equ 55
    IDC_CHECK_MAIN_GUI_CANCEL_EXEC      equ 56
    IDC_CHECK_MAIN_TERM                 equ 57
    IDC_CHECK_MAIN_COPY_AFTER           equ 58
    IDC_CHECK_MAIN_ASK_COPY_AFTER       equ 59
    IDC_CHECK_MAIN_GUI_COPY_AFTER       equ 60
    IDC_PROMPT_MAIN_TIMEOUT             equ 61
    IDC_EDIT_MAIN_TIMEOUT               equ 62
    IDC_SPINBOX_MAIN_TIMEOUT            equ 63
    IDC_CHECK_MAIN_COPY_BEFORE          equ 64
    IDC_CHECK_MAIN_ASK_COPY_BEFORE      equ 65
    IDC_CHECK_MAIN_GUI_COPY_BEFORE      equ 66
    ;
    IDC_BUTTON_ABOUT                    equ 67
    IDC_BUTTON_SAVE                     equ 68
    IDC_BUTTON_CLOSE                    equ 69
    ;
    IDC_BUTTON_CLADEBUG                 equ 70
    IDC_BUTTON_DEBUGSPY                 equ 71
    ;
    EQU_LAST_CONTROL                    equ IDC_BUTTON_CLOSE
    ;
    IDGROUP3                            equ 500    ;About Window equates
    IDSTATIC12                          equ 501
    IDC_BTN6                            equ 502
    ;
    IDM_COMPUTERNAME                    equ 200
    IDM_DATE                            equ 201
    IDM_DAY                             equ 202
    IDM_DRIVE                           equ 203
    IDM_ERRORCODE                       equ 204
    IDM_ERRORTEXT                       equ 205
    IDM_HOUR                            equ 206
    IDM_MINUTE                          equ 207
    IDM_MONTH                           equ 208
    IDM_MYDOCUMENTS                     equ 209
    IDM_PROCESSID                       equ 210
    IDM_PROGRAMDIR                      equ 211
    IDM_PROGRAMFILES                    equ 212
    IDM_PROGRAMNAME                     equ 213
    IDM_SECOND                          equ 214
    IDM_SHARE                           equ 215
    IDM_SYSTEMDIR                       equ 216
    IDM_SYSTEMDRIVE                     equ 217
    IDM_TEMPDIR                         equ 218
    IDM_TIME                            equ 219
    IDM_USERNAME                        equ 220
    IDM_WEEKDAY                         equ 221
    IDM_WINDIR                          equ 222
    IDM_YEAR                            equ 223
    IDM_23                              equ 224
    EQU_HOTKEY                          equ 225
    ;
.data
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;UNITIALIZED DATA SECTION
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
    ;
    szToolTipClass              db "tooltips_class32",0
    ;
    szExtentionDotExe           db ".exe",0
    ;
    szToolTxt00                 db "Press CTRL+SPACEBAR to PopUp Paste-In-Macros menu. Use semicolon "";"" for multiple wildcard: *.doc; *.xls",0
    szToolTxt0                  db "Press CTRL+SPACEBAR to PopUp Paste-In-Macros menu",0
    szToolTxt1                  db "Process Subfolders",0
    szToolTxt2                  db "Overwrite if file exists",0
    szToolTxt3                  db "For debugging: do NOT HIDE the window of the 'Execute' application and updates %PROCESSID% macro value."
                                db " If an instance of the 'Execute' is already running then only updates %PROCESSID% macro value",0 ;TODO if found and "ask"=true then do not ask.
    szToolTxt4                  db "Seconds to terminate the instance(TERM=1) or seconds to wait (Show=1) after running the 'Execute' application",0
    szToolTxt5                  db "Unhide and bring window to top if the 'Execute' application is running",0
    szToolTxt6                  db "Terminate the 'Execute' instances if running (Ignored if Unhide/Show is Checked)",0
    szToolTxt7                  db "Shows a Message asking to Copy the files from Source to Target",0
    szToolTxt8                  db "Shows a Message asking to 'Execute' the Application",0
    szToolTxt9                  db "Add a Cancel button to the 'Ask before' Copy message to open this window",0
    szToolTxtA                  db "Add a Cancel button to the 'Ask before' Execute message to open this window",0
    szToolTxtB                  db "Copy files after 'Execute' the application",0
    szToolTxtC                  db "Copy files before 'Execute' the application",0
    szToolTxtD                  db "Setting is saved in a .ini file with the same name as the executable file"
                                db " e.g.: RunHide.exe = Run Hide.ini, AnyName.exe = AnyName.ini, etc."
                                db " thus allow having multiple environment for developing or whatever",0
    szToolTxtE                  db "Press CTRL key when execute RunHide to show this window setting",0
    ;
    szWebPage                   db 'http://www.carabez.com',0
    PROP_ORIGINAL_FONT          db '_Hyperlink_Original_Font_',0
    PROP_ORIGINAL_PROC          db '_Hyperlink_Original_Proc_',0
    PROP_UNDERLINE_FONT         db '_Hyperlink_Underline_Font_',0
    szFileVersion               db "\\StringFileInfo\\040904E4\\FileVersion",0
    szVer                       db " - Version: ",0
    ;
    szSecPrevious               db "PREVIOUS",0
    szSecMainOptions            db "OPTIONS",0
    szKeyCopySource             db "CopySource",0
    szKeyCopyTarget             db "CopyTarget",0
    szKeyCopySubFolders         db "CopySubFolders",0
    szKeyCopyOverWrite          db "CopyOverwrite",0
    szKeyProgram                db "Program",0
    szKeyParameters             db "Parameters",0
    szKeyWorkingDir             db "WorkingDir",0
    szKeyUnhide                 db "Unhide",0
    szKeyShow                   db "Show",0
    szKeyAskExecute             db "AskExecute",0
    szKeyGuiOnCancelExecute     db "GuiOnCancelExecute",0
    szKeyTerm                   db "Term",0
    szKeyCopyAfterExecute       db "CopyAfterExecute",0
    szKeyAskCopyAfter           db "AskCopy",0
    szKeyGuiOnCancelCopyAfter   db "GuiOnCancelCopy",0
    szKeyTimeout                db "Timeout",0
    szKeyCopyBeforeExecute      db "CopyBeforeExecute",0
    szKeyAskCopyBefore          db "AskCopyBefore",0
    szKeyGuiOnCancelCopyBefore  db "GuiOnCancelCopyBefore",0
    ;
    szTODO                      db "TODO",0
    szCurrentVersion            db "Software\Microsoft\Windows\CurrentVersion",0
    szProgramFilesDir           db "ProgramFilesDir",0
    szTwoDigitFormat            db "%02d",0                 ; To format '1' into '01:'
    szNewLine                   db 0AH,0 ;0AH,0CH,
    ;szNoFilter                 db "Program Files",0,"*.exe",0,"All files",0,"*.*",0,0
    szCOMPUTERNAME              db "%COMPUTERNAME%",0         ;Computer name of the current system
    szDATE                      db "%DATE%",0                 ;Current date (in the format YYYYMMDD)
    szDAY                       db "%DAY%",0                  ;Day of the month (1-31)
    szDRIVE                     db "%DRIVE%",0                ;Drive that is redirected to ShareAddress (D  db  C  db  etc)
    szERRORCODE                 db "%ERRORCODE%",0            ;Last Error Code Returned by  GetLastError
    szERRORTEXT                 db "%ERRORTEXT%",0            ;Last Error Text Returned by  GetLastError
    szHOUR                      db "%HOUR%",0                 ;LocalTime Hour 0-24
    szMINUTE                    db "%MINUTE%",0               ;LocalTime Hour Minute 00-59
    szMONTH                     db "%MONTH%",0                ;Month of the Year (1-12)
    szMYDOCUMENTS               db "%MYDOCUMENTS%",0          ;User ID My Documents.
    szPROCESSID                 db "%PROCESSID%",0            ;-p The debuggee's process ID
    szPROGRAMDIR                db "%PROGRAMDIR%",0           ;Path of the Program Working directory
    szPROGRAMFILES              db "%PROGRAMFILES%",0         ;Path of the Program Files Directory
    szPROGRAMNAME               db "%PROGRAMNAME%",0          ;Name of the current executable file  usually = MapDrive
    szSECOND                    db "%SECOND%",0               ;LocalTime Second 00-59
    szSHARE                     db "%SHARE%",0                ;Full Share Address
    szSYSTEMDIR                 db "%SYSTEMDIR%",0            ;Path of the Windows system directory
    szSYSTEMDRIVE               db "%SYSTEMDRIVE%",0          ;Default Root Drive usually = C:\
    szTEMPDIR                   db "%TEMPDIR%",0              ;Default Temp Directory
    szTIME                      db "%TIME%",0                 ;Current time (in the format HHMMSS)
    szUSERNAME                  db "%USERNAME%",0             ;Current user's Windows user ID
    szWEEKDAY                   db "%WEEKDAY%",0              ;Days since Sunday (1 Ц 7)
    szWINDIR                    db "%WINDIR%",0               ;Path of the Windows directory
    szYEAR                      db "%YEAR%",0                 ;Current year
    ;
    szSpace                     db " ",0
    szBracketOpen               db " [",0
    szTimeOutOpen               db " [TIMEOUT=",0
    szBracketClose              db "]",0
    szAskForCopyTitlePre        db "PREVIOUS TASK: Please confirm",0
    szAskForCopyTitleMain       db "MAIN TASK: Please confirm",0
    szAskForCopyText1           db "Copy file(s)?",10,10,0
    szAskForCopyText2           db 10,10,"To",10,10,0
    ;
    szAskForExecTitlePre        db "PREVIOUS TASK: Please confirm",0
    szAskForExecTitleMain       db "MAIN TASK: Please confirm",0
    szAskForExecExecute         db "Execute application in HIDDEN mode?",10,10,0
    szAskForExecShow            db "Execute and SHOW the application?",10,10,0
    szAskForExecTerminate       db "TERMINATE application?",10,10,0
    szAskForExecUnhide          db "UNHIDE application?",10,10,0
    szAskForExecParameters      db 10,10,"Parameters:",10,10,0
    szAskForExecWorkingDir      db 10,10,"Working Directory:",10,10,0
    ;
    szAskForFileTitle           db "Please Confirm",0
    szAskForFileText            db "Is it a request for a folder?",0
    ;
    szFileDialogTitle           db "Selecting a file",0
    szFileDialogFolder          db "Please select Target folder",0
    szFileDialogFilter          db "All files",0,"*.*",0,0
    szFileDialogNoFilter        db "All files",0,"*.*",0,0
    ;
    szCmdTerm                   db "TERM",0
    szCmdUnhide                 db "UNHIDE",0
    szCmdTimeout                db "TIMEOUT",0
    szCmdShow                   db "SHOW",0
    szTrue                      db "TRUE",0
    szExtention                 db ".exe",0
    szErrorSavingIniFileTitle   db "Error Saving Ini File",0
    szErrorExecuteTitle         db "Error when executing",0
    szUsageTitle                db "RunHide /?",0
    szUsageText                 db "SYNTAX:",10,10
                                db "RunHide [program] [parameters] [TIMEOUT=seconds] [TERM] [UNHIDE] [SHOW=1]",10,10
                                ;db "RunHide [program] [parameters] [TIMEOUT=seconds] [TERM] [UNHIDE] [P=ProcessID] [SHOW=1]",10,10
                                db "TIMEOUT",10,9,"If after the specified number of seconds,",10,9
                                db      "the launched program has not terminated,",10,9
                                db      "it will be forcefully terminated.",10
                                db "TERM",10,9,"If you did not specify a TIMEOUT value,",10,9
                                db      "but you find the program is still running hidden,",10,9
                                db      "you may terminate the program by specifying the TERM option,",10,9
                                db      "or you may unhide the program by specifying the UNHIDE option.",10,9
                                db      "For example:",10,10,9
                                db      "RunHide MyProgram TERM",10,10,9
                                db      "This will kill all instances of MYPROGRAM.EXE",10
                                db "UNHIDE",10,9,"If program is running: Unhide and bring window to top.",10,10
                                ;db "P",10,9,"A Process ID in decimal or hexadecimal format.",10,10
                                db "SHOW",10,9,"For debugging: when True or 1 shows the program to execute.",10,10
                                db  "OPTIONAL INI FILE",10,9,"Options and parameters can be taken from RunHide.ini",10,10,9
                                db      "Do you need more info about RunHide.ini?",0
    szIniTitle                  db "The RunHide initialization file example (runhide.ini).",0
    szIniText                   db "[OPTIONS]",10,9
                                db ";Section where the parameters will be localized.",10
                                db "Program=notepad",10,9
                                db ";The Program to execute.",10
                                db "Parameters=readme.txt",10,9
                                db ";The Parameter(s) to pass to the program to execute.",10
                                db "Timeout=0",10,9
                                db ";Seconds to wait until kill the executed program. When zero this option is disable.",10
                                db "Term=0",10,9
                                db ";when term=1 the application kills any instance of the previously executed program.",10
                                db "Unhide=0",10,9
                                db ";when Unhide=1 any instance of the previously executed program is unhidden.",10
                                db "Show=0",10,9
                                db ";when Show=1 the application DO NOT hide the Program to execute.",10,10,10
                                db ";Note 1:",10,";",9,"The ini file is formed from application name; so if you rename RunHide.exe",10
                                db ";",9,"to, by example, Runftp.exe the expected ini file will be Runftp.ini",10
                                db ";Note 2:",10,";",9,"Command line argument are ignored when an ini file is found.",10
                                db ";Note 3:",10,";",9,"To open the Graphical user interface window press and hold CTRL key then RUN RunHide.exe ",10
                                db ";Note 4:",10,";",9,"To copy the entire contents of THIS (or any) message box Press CTRL+C",10,10,9
                                db "Do you want to Enter in GRAPHICAL mode now?",0
                                ;
    ;
    ;align 4 ;SHGetKnownFolderPath SHELL32 v6.0 >= Vista
    ;FOLDERID_Documents          GUID  <0FDD39AD0h,238Fh,46AFh,<0ADh,0B4h,06Ch,85h,48h,03h,69h,0C7h>> ;'{FDD39AD0-238F-46AF-ADB4-6C85480369C7}' name GUID <l,w1,w2,<b1,b2,b3,b4,b5,b6,b7,b8>>
    
    
;    typedef enum {
;      KF_FLAG_DEFAULT = 0x00000000,
;      KF_FLAG_FORCE_APP_DATA_REDIRECTION = 0x00080000,
;      KF_FLAG_RETURN_FILTER_REDIRECTION_TARGET = 0x00040000,
;      KF_FLAG_FORCE_PACKAGE_REDIRECTION = 0x00020000,
;      KF_FLAG_NO_PACKAGE_REDIRECTION = 0x00010000,
;      KF_FLAG_FORCE_APPCONTAINER_REDIRECTION = 0x00020000,
;      KF_FLAG_NO_APPCONTAINER_REDIRECTION = 0x00010000,
;      KF_FLAG_CREATE = 0x00008000,
;      KF_FLAG_DONT_VERIFY = 0x00004000,
;      KF_FLAG_DONT_UNEXPAND = 0x00002000,
;      KF_FLAG_NO_ALIAS = 0x00001000,
;      KF_FLAG_INIT = 0x00000800,
;      KF_FLAG_DEFAULT_PATH = 0x00000400,
;      KF_FLAG_NOT_PARENT_RELATIVE = 0x00000200,
;      KF_FLAG_SIMPLE_IDLIST = 0x00000100,
;      KF_FLAG_ALIAS_ONLY = 0x80000000
;    } KNOWN_FOLDER_FLAG;
                                                                         
    align 16
    abntbl                      db 2,0,0,0,0,0,0,0,0,1,1,0,0,1,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 1,0,3,0,0,0,0,0,0,0,0,0,1,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                db 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
                                ; 0 = OK char
                                ; 1 = delimiting characters   tab LF CR space ","
                                ; 2 = ASCII zero    This should not be changed in the table
                                ; 3 = quotation     This should not be changed in the table
                                ;
    ;
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
.data?
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;UNITIALIZED DATA SECTION
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
    align 4
    hInstance                   dd ?
    hIcon                       dd ?
    hIconClaDebug               dd ?
    bIsClarionDebugRunning      dd ?
    hDesktopWin                 dd ?
    ;
    hPopupMenu                  dd ?                     ;Handle of PopUpMenu
    hOwnerPopMenu               dd ?                     ;Active Windows'Handle target of PopUpMenu
    hFocus                      dd ?                     ;To get Handle of control with focus
    ;
    ti                          TOOLINFO<>
    icex                        INITCOMMONCONTROLSEX <> ;structure for Controls
    hToolTip                    dd ?
    ;
    dErrorCode                  dd ?
      ;
    hWinMain                    dd ?
    ;
    ;Version
    dDummy                      dd ?
    pMemory                     dd ?
    dwVersionLen                dd ?
    hBmp                        dd ?
    szBuffer                    db MAX_PATH dup(?)
    szOutputDebugString         db 4096 dup(?)
    ;
    ;
    pParameters                 dd ?
    szIniFile                   db 4096 dup(?)
    szSolved1                   db 4096 dup(?)
    szSolved2                   db 4096 dup(?)
    szSolved3                   db 4096 dup(?)
    szRunOptions                db MAX_PATH dup(?)
    szCmdCompare                db MAX_PATH dup(?)
    ;wfd                         WIN32_FIND_DATA <?>
    ;
    szProgramPath               db MAX_PATH dup(?)
    szProgramName               db MAX_PATH dup(?)
    szProgramNameDotIni         db MAX_PATH dup(?)
    szWindowsDir                db MAX_PATH dup(?)
    ;
    szMessageBoxTitle           db MAX_PATH dup(?)
    szMessageBoxText            db 4096 dup(?)
    ;
    hProcessID_ForMacro         dd ?
    ;
    align 4
    szPreCopySource             db 4096 dup(?) ;MAX_PATH dup(?)
    dPreCopySubFolders          dd ?
    szPreCopyTarget             db 4096 dup(?) ;MAX_PATH dup(?)
    dPreCopyOverwrite           dd ?
    szPreCmdRunProgram          db 4096 dup(?)
    szPreCmdRunParameters       db 4096 dup(?)
    szPreCmdRunWorkingDir       db 4096 dup(?)
    dPreUnhide                  dd ?
    dPreShow                    dd ?
    dPreAskExecute              dd ?
    dPreGuiOnCancelExecute      dd ?
    dPreTerm                    dd ?
    dPreCopyAfterExecute        dd ?
    dPreAskCopyAfterExecute     dd ?
    dPreGuiOnCancelCopyAfter    dd ?
    dPreTimeout                 dd ?
    dPreCopyBeforeExecute       dd ?
    dPreAskCopyBeforeExecute    dd ?
    dPreGuiOnCancelCopyBefore   dd ?
    ;
    align 4
    szCopySource                db 4096 dup(?) ;MAX_PATH dup(?)
    dCopySubFolders             dd ?
    szCopyTarget                db 4096 dup(?) ;MAX_PATH dup(?)
    dCopyOverwrite              dd ?
    szCmdRunProgram             db 4096 dup(?)
    szCmdRunParameters          db 4096 dup(?)
    szCmdRunWorkingDir          db 4096 dup(?)
    dUnhide                     dd ?
    dShow                       dd ?
    dAskExecute                 dd ?
    dGuiOnCancelExecute         dd ?
    dTerm                       dd ?
    dCopyAfterExecute           dd ?
    dAskCopyAfterExecute        dd ?
    dGuiOnCancelCopyAfter       dd ?
    dTimeout                    dd ?
    dCopyBeforeExecute          dd ?
    dAskCopyBeforeExecute       dd ?
    dGuiOnCancelCopyBefore      dd ?
    ;
    ACTIVESIDC                  struct
        hIDC                        dd ?
        hProc                       dd ?
    ACTIVESIDC                  ends
    ActiveIDC                   ACTIVESIDC 10 dup(<?>)
    ;
    szFullPathApp               db 4096 dup(?)
    szFullPathLocal             db 4096 dup(?)
    szFullPathWorkingDir        db 4096 dup(?)
    szCommandLineParameters     db 4096 dup(?)
    ;
    BtnStruct                   XXBUTTON <?>
    ;hBmp                        dd ?
    hBmpNoiseNormal             dd ?
    hBmpNoiseHover              dd ?
    DrawingRec                  RECT<>
    WinMainRec                  RECT<>
    UpdateRec                   RECT<>
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
.code
start:
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ; Code Entry Point
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
    ;

    Invoke RunCmdAtMainEntry
    .if eax
        call main
    .endif
    xor eax,eax
    Invoke ExitProcess,eax
        ;
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
RunCmdAtMainEntry               proc
                                ;
    ;
    Invoke GetModuleHandle,NULL
    mov hInstance,eax
    mov bIsClarionDebugRunning,FALSE
    Invoke LoadIcon,hInstance,500
    mov hIcon,eax
    ;Invoke LoadIcon,hInstance,501
    mov hIconClaDebug,eax
    ;
    Invoke InitCommonControls
    ;
    Invoke GetDesktopWindow
    mov hDesktopWin,eax
    Invoke GetModuleFileName,hInstance,addr szBuffer,sizeof szBuffer
    ;Get Application Path and Name
    Invoke GetCmdLine,0,addr szProgramPath
    Invoke lstrcmpi,addr szProgramPath,addr szBuffer
    .if eax!=NULL   ;No Equals, that means tha was called by a win16 app.
        ;Check for the cmd arg.
        Invoke GetCmdLine,1,addr szProgramPath
        ;1=Sucess,2=No Argument,3=NoMatch
        .if eax==2
            ;Invoke MessageBox,NULL,addr szBuffer,addr szProgramPath,NULL
            Invoke lstrcpy,addr szProgramPath,addr szBuffer
        .endif
    .endif
    Invoke lstrcpy,addr szIniFile,addr szProgramPath
    ;Invoke GetCmdLine,0,addr szIniFile


    ;
    ;Change Name .exe to .ini
    Invoke lstrlen, addr szIniFile
    lea ecx,szIniFile
    mov byte ptr [ecx+eax-1],'i'
    mov byte ptr [ecx+eax-2],'n'
    mov byte ptr [ecx+eax-3],'i'
    ;
    ;********** Get Only the ProgramName.ini
    Invoke lstrlen, addr szIniFile
    lea ecx,szIniFile
    add eax,ecx
    @@:
    dec eax                       ;Decrement the string
    cmp byte ptr [eax],'\'
    jne @B
    inc eax
    Invoke lstrcpy,addr szProgramName,eax
    Invoke lstrcpy,addr szProgramNameDotIni,addr szProgramName
    ;
    Invoke lstrlen, addr szProgramName
    lea ecx,szProgramName
    add eax,ecx
    sub eax,4
    mov byte ptr [eax],00H                            ;Remove .ini
    ;
    xor eax,eax
    ;
    mov hProcessID_ForMacro,eax
    ;
    mov dPreShow,eax
    mov dPreCopyAfterExecute,eax
    mov dPreTimeout,eax
    mov dPreUnhide,eax
    mov dPreTerm,eax
    mov dPreAskCopyAfterExecute,eax
    mov dPreAskExecute,eax
    mov dPreGuiOnCancelCopyAfter,eax
    mov dPreGuiOnCancelExecute,eax
    ;
    mov dShow,eax
    mov dCopyAfterExecute,eax
    mov dTimeout,eax
    mov dUnhide,eax
    mov dTerm,eax
    mov dAskCopyAfterExecute,eax
    mov dAskExecute,eax
    mov dGuiOnCancelCopyAfter,eax
    mov dGuiOnCancelExecute,eax
    ;
    mov dUnhide,FALSE
    mov dShow,SW_HIDE
    ;
    Invoke FileExist,addr szIniFile
    .if eax == 0                    ;Ini File DO NOT Exists
        CmdLineHasPriorityOverIniFile:
        Invoke CheckCommandLineArgs
        .if eax==0                      ;FALSE=Show Usage
            ;
            ;Check if CTRL KEY is Pressed
            Invoke GetKeyState,VK_CONTROL ;VK_MENU ;VK_SHIFT ;VK_CONTROL ;VK_CONTROL = RIGHT_CTRL_PRESSED || LEFT_CTRL_PRESSED
            .if (ax > 1) ;If the high-order bit is 1, the key is down; otherwise, it is up.
                 mov eax,TRUE
                 jmp ExitAndCallMain
            .endif
            ;
            ;
            Invoke MessageBox,NULL,addr szUsageText,addr szUsageTitle,MB_YESNO+MB_ICONQUESTION
            .if eax == IDYES
                Invoke MessageBox,NULL,addr szIniText,addr szIniTitle,MB_YESNO+MB_ICONQUESTION
                .if eax == IDYES
                    mov eax,TRUE
                    jmp ExitAndCallMain
                .endif
            .endif
            jmp ExitRunCmd
        .endif
        ;
        ;Invoke MessageBox,NULL,OFFSET szCommandLineParameters,OFFSET szCmdRunProgram,NULL
        ;Invoke MessageBox,NULL,addr szCommandLineParameters,addr szCmdRunProgram,NULL
        ;Invoke Sleep,1000
        Invoke ProccessParameters,addr szCmdRunProgram,addr szCommandLineParameters,NULL,dTerm,dUnhide,dShow,dTimeout
        mov hProcessID_ForMacro,eax
        ;Invoke MessageBox,NULL,OFFSET szCommandLineParameters,OFFSET szCmdRunProgram,NULL
        ;
    .else                           ;Ini File exists
        ;
        ;Check if CTRL KEY is Pressed
        Invoke GetKeyState,VK_CONTROL ;VK_MENU ;VK_SHIFT ;VK_CONTROL ;VK_CONTROL = RIGHT_CTRL_PRESSED || LEFT_CTRL_PRESSED
        .if (ax > 1) ;If the high-order bit is 1, the key is down; otherwise, it is up.
            mov eax,TRUE
            jmp ExitAndCallMain
        .endif
        ;CmdLine parameters has priority over Ini File:
        ;Invoke MessageBox,NULL,addr szCmdRunParameters,addr szCmdRunProgram,NULL
        Invoke GetCmdLine,1,addr szCmdRunParameters ;1=Sucess,2=No Argument,3=NoMatch
        .if eax==1
            jmp CmdLineHasPriorityOverIniFile
        .endif
        ;
        ;**********************************************
        ;**********************************************
        ; RUNHIDE PREVIOUS TASK
        ;**********************************************
        ;**********************************************
        Invoke ReadIniFile,NULL
        ;
        ; Update hProcessID_ForMacro
        Invoke lstrlen,addr szPreCmdRunProgram
        .if eax!=0
            Invoke SolveMacros,addr szPreCmdRunProgram,addr szSolved1 ;source, target
            ;Invoke OutputDebugString,addr szSolved1
            Invoke FindApplicationProcessID,addr szSolved1,-1,NULL
            mov hProcessID_ForMacro,eax
        .endif
        ;
        ;
        Invoke lstrlen,addr szPreCopySource
        .if eax==0
            jmp @F
        .endif
        .if dPreCopyBeforeExecute
            .if dPreAskCopyBeforeExecute
                Invoke lstrcpy,addr szMessageBoxTitle,addr szAskForCopyTitlePre
                ;
                Invoke lstrcpy,addr szMessageBoxText,addr szAskForCopyText1
                Invoke SolveMacros,addr szPreCopySource,addr szSolved1 ;source, target
                Invoke lstrcat,addr szMessageBoxText,addr szSolved1
                Invoke lstrcat,addr szMessageBoxText,addr szAskForCopyText2
                Invoke SolveMacros,addr szPreCopyTarget,addr szSolved2
                Invoke lstrcat,addr szMessageBoxText,addr szSolved2
                ;
                Invoke Sleep,50
                Invoke ForceWindowToTop,hDesktopWin
                ;
                .if dPreGuiOnCancelCopyBefore
                    mov ecx,MB_YESNOCANCEL+MB_ICONQUESTION+MB_DEFBUTTON1+MB_TOPMOST+MB_SYSTEMMODAL
                .else
                    mov ecx,MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1+MB_TOPMOST+MB_SYSTEMMODAL
                .endif
                Invoke MessageBox,hDesktopWin,addr szMessageBoxText,addr szMessageBoxTitle,ecx
                .if eax == IDNO
                    jmp @F
                .elseif eax == IDCANCEL
                    mov eax,TRUE
                    jmp ExitAndCallMain
                .endif
            .endif
            Invoke SolveMacros,addr szPreCopySource,addr szSolved1 ;source, target
            Invoke SolveMacros,addr szPreCopyTarget,addr szSolved2
            Invoke CopyFiles,addr szSolved1,addr szSolved2,dPreCopySubFolders,dPreCopyOverwrite
        .endif
        @@:
        Invoke lstrlen,addr szPreCmdRunProgram
        .if eax==0
            jmp @F
        .endif
        ;
        .if dPreAskExecute
            ;
            Invoke lstrcpy,addr szMessageBoxTitle,addr szAskForExecTitlePre
            ;
            .if dPreShow == TRUE ;Show has priority over Unhide and Term
                Invoke lstrcpy,addr szMessageBoxText,addr szAskForExecShow
                .if dPreTimeout
                    Invoke lstrcat,addr szMessageBoxTitle,addr szTimeOutOpen
                    Invoke DwordToAsccii,dPreTimeout,addr szBuffer
                    Invoke lstrcat,addr szMessageBoxTitle,addr szBuffer
                    Invoke lstrcat,addr szMessageBoxTitle,addr szBracketClose
                .endif
            .else
                .if dPreUnhide == TRUE ;Unhide has priority over Term
                    Invoke lstrcpy,addr szMessageBoxText,addr szAskForExecUnhide
                .elseif dPreTerm == TRUE
                    Invoke lstrcpy,addr szMessageBoxText,addr szAskForExecTerminate
                .else
                    Invoke lstrcpy,addr szMessageBoxText,addr szAskForExecExecute
                    .if dPreTimeout
                        Invoke lstrcat,addr szMessageBoxTitle,addr szTimeOutOpen
                        Invoke DwordToAsccii,dPreTimeout,addr szBuffer
                        Invoke lstrcat,addr szMessageBoxTitle,addr szBuffer
                        Invoke lstrcat,addr szMessageBoxTitle,addr szBracketClose
                    .endif
                .endif
            .endif
            Invoke SolveMacros,addr szPreCmdRunProgram,addr szSolved1 ;source, target
            Invoke lstrcat,addr szMessageBoxText,addr szSolved1 ;szCmdRunProgram
            Invoke lstrlen,addr szPreCmdRunParameters
            .if eax!=0
                Invoke SolveMacros,addr szPreCmdRunParameters,addr szSolved2
                Invoke lstrcat,addr szMessageBoxText,addr szAskForExecParameters
                Invoke lstrcat,addr szMessageBoxText,addr szSolved2 ;szCmdRunParameters
            .endif
            ;
            Invoke lstrlen, addr szPreCmdRunWorkingDir
            .if eax!=0
                Invoke SolveMacros,addr szPreCmdRunWorkingDir,addr szSolved3
                Invoke lstrcat,addr szMessageBoxText,addr szAskForExecWorkingDir
                Invoke lstrcat,addr szMessageBoxText,addr szSolved3 ;szCmdRunWorkingDir
            .endif
            ;
            Invoke Sleep,50
            Invoke ForceWindowToTop,hDesktopWin
            ;
            .if dPreGuiOnCancelExecute
                mov ecx,MB_YESNOCANCEL+MB_ICONQUESTION+MB_DEFBUTTON1+MB_TOPMOST+MB_SYSTEMMODAL ;MB_TASKMODAL ;MB_SYSTEMMODAL
            .else
                mov ecx,MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1+MB_TOPMOST+MB_SYSTEMMODAL ;MB_TASKMODAL ;MB_SYSTEMMODAL
            .endif
            Invoke MessageBox,hDesktopWin,addr szMessageBoxText,addr szMessageBoxTitle,ecx
            .if eax == IDNO
                jmp @F
            .elseif eax == IDCANCEL
                mov eax,TRUE
                jmp ExitAndCallMain
            .endif
        .endif
        ;**********************************************
        ; Execute Previous
        Invoke SolveMacros,addr szPreCmdRunProgram,addr szSolved1 ;source, target
        Invoke SolveMacros,addr szPreCmdRunParameters,addr szSolved2
        Invoke SolveMacros,addr szPreCmdRunWorkingDir,addr szSolved3
        Invoke ProccessParameters,addr szSolved1,addr szSolved2,addr szSolved3,dPreTerm,dPreUnhide,dPreShow,dPreTimeout
        mov hProcessID_ForMacro,eax
        ;
        ;**********************************************
        @@:
        Invoke lstrlen,addr szPreCopySource
        .if eax==0
            jmp @F
        .endif
        .if dPreCopyAfterExecute
            .if dPreAskCopyAfterExecute
                Invoke lstrcpy,addr szMessageBoxTitle,addr szAskForCopyTitlePre
                ;
                Invoke lstrcpy,addr szMessageBoxText,addr szAskForCopyText1
                Invoke SolveMacros,addr szPreCopySource,addr szSolved1 ;source, target
                Invoke lstrcat,addr szMessageBoxText,addr szSolved1 ;szCopySource
                Invoke lstrcat,addr szMessageBoxText,addr szAskForCopyText2
                Invoke SolveMacros,addr szPreCopyTarget,addr szSolved2
                Invoke lstrcat,addr szMessageBoxText,addr szSolved2 ;szCopyTarget
                ;
                Invoke Sleep,50
                Invoke ForceWindowToTop,hDesktopWin
                ;
                .if dPreGuiOnCancelCopyAfter
                    mov ecx,MB_YESNOCANCEL+MB_ICONQUESTION+MB_DEFBUTTON1+MB_TOPMOST+MB_SYSTEMMODAL
                .else
                    mov ecx,MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1+MB_TOPMOST+MB_SYSTEMMODAL
                .endif
                Invoke MessageBox,hDesktopWin,addr szMessageBoxText,addr szMessageBoxTitle,ecx
                .if eax == IDNO
                    jmp @F
                .elseif eax == IDCANCEL
                    mov eax,TRUE
                    jmp ExitAndCallMain
                .endif
            .endif
            Invoke SolveMacros,addr szPreCopySource,addr szSolved1 ;source, target
            Invoke SolveMacros,addr szPreCopyTarget,addr szSolved2
            Invoke CopyFiles,addr szSolved1,addr szSolved2,dCopySubFolders,dCopyOverwrite
        .endif
        @@:
        ;**********************************************
        ; RUNHIDE MAIN TASK
        ;**********************************************
        Invoke lstrlen,addr szCopySource
        .if eax==0
            jmp @F
        .endif
        .if dCopyBeforeExecute
            .if dAskCopyBeforeExecute
                Invoke lstrcpy,addr szMessageBoxTitle,addr szAskForCopyTitleMain
                ;
                Invoke lstrcpy,addr szMessageBoxText,addr szAskForCopyText1
                Invoke SolveMacros,addr szCopySource,addr szSolved1 ;source, target
                Invoke lstrcat,addr szMessageBoxText,addr szSolved1 ;szCopySource
                Invoke lstrcat,addr szMessageBoxText,addr szAskForCopyText2
                Invoke SolveMacros,addr szCopyTarget,addr szSolved2
                Invoke lstrcat,addr szMessageBoxText,addr szSolved2 ;szCopyTarget
                ;
                Invoke Sleep,50
                Invoke ForceWindowToTop,hDesktopWin
                ;
                .if dGuiOnCancelCopyAfter
                    mov ecx,MB_YESNOCANCEL+MB_ICONQUESTION+MB_DEFBUTTON1+MB_TOPMOST+MB_SYSTEMMODAL
                .else
                    mov ecx,MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1+MB_TOPMOST+MB_SYSTEMMODAL
                .endif
                Invoke MessageBox,hDesktopWin,addr szMessageBoxText,addr szMessageBoxTitle,ecx
                .if eax == IDNO
                    jmp @F
                .elseif eax == IDCANCEL
                    mov eax,TRUE
                    jmp ExitAndCallMain
                .endif
            .endif
            Invoke SolveMacros,addr szCopySource,addr szSolved1 ;source, target
            Invoke SolveMacros,addr szCopyTarget,addr szSolved2
            Invoke CopyFiles,addr szSolved1,addr szSolved2,dCopySubFolders,dCopyOverwrite
        .endif
        @@:
        Invoke lstrlen,addr szCmdRunProgram
        .if eax==0
            jmp @F
        .endif
        .if dAskExecute
            ;
            Invoke lstrcpy,addr szMessageBoxTitle,addr szAskForExecTitleMain
            ;
            .if dShow == TRUE ;Show has priority over Unhide and Term
                Invoke lstrcpy,addr szMessageBoxText,addr szAskForExecShow
                .if dTimeout
                    Invoke lstrcat,addr szMessageBoxTitle,addr szTimeOutOpen
                    Invoke DwordToAsccii,dTimeout,addr szBuffer
                    Invoke lstrcat,addr szMessageBoxTitle,addr szBuffer
                    Invoke lstrcat,addr szMessageBoxTitle,addr szBracketClose
                .endif
            .else
                .if dUnhide == TRUE ;Unhide has priority over Term
                    Invoke lstrcpy,addr szMessageBoxText,addr szAskForExecUnhide
                .elseif dTerm == TRUE
                    Invoke lstrcpy,addr szMessageBoxText,addr szAskForExecTerminate
                .else
                    Invoke lstrcpy,addr szMessageBoxText,addr szAskForExecExecute
                    .if dTimeout
                        Invoke lstrcat,addr szMessageBoxTitle,addr szTimeOutOpen
                        Invoke DwordToAsccii,dTimeout,addr szBuffer
                        Invoke lstrcat,addr szMessageBoxTitle,addr szBuffer
                        Invoke lstrcat,addr szMessageBoxTitle,addr szBracketClose
                    .endif
                .endif
            .endif
            Invoke SolveMacros,addr szCmdRunProgram,addr szSolved1 ;source, target
            Invoke lstrcat,addr szMessageBoxText,addr szSolved1 ;szCmdRunProgram
            ;
            Invoke lstrlen,addr szCmdRunParameters
            .if eax!=0
                Invoke SolveMacros,addr szCmdRunParameters,addr szSolved2
                Invoke lstrcat,addr szMessageBoxText,addr szAskForExecParameters
                Invoke lstrcat,addr szMessageBoxText,addr szSolved2 ;szCmdRunProgram
            .endif
            ;
            Invoke lstrlen, addr szCmdRunWorkingDir
            .if eax!=0
                Invoke SolveMacros,addr szCmdRunWorkingDir,addr szSolved3
                Invoke lstrcat,addr szMessageBoxText,addr szAskForExecWorkingDir
                Invoke lstrcat,addr szMessageBoxText,addr szSolved3 ;szCmdRunWorkingDir
            .endif
            ;
            Invoke Sleep,50
            Invoke ForceWindowToTop,hDesktopWin
            ;
            .if dGuiOnCancelExecute
                mov ecx,MB_YESNOCANCEL+MB_ICONQUESTION+MB_DEFBUTTON1+MB_TOPMOST+MB_SYSTEMMODAL
            .else
                mov ecx,MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1+MB_TOPMOST+MB_SYSTEMMODAL
            .endif
            Invoke MessageBox,hDesktopWin,addr szMessageBoxText,addr szMessageBoxTitle,ecx
            .if eax == IDNO
                jmp @F
            .elseif eax == IDCANCEL
                mov eax,TRUE
                jmp ExitAndCallMain
            .endif
        .endif
        ;**********************************************
        ; Execute Main
        Invoke SolveMacros,addr szCmdRunProgram,addr szSolved1 ;source, target
        Invoke SolveMacros,addr szCmdRunParameters,addr szSolved2
        Invoke SolveMacros,addr szCmdRunWorkingDir,addr szSolved3
        ;Invoke MessageBox,NULL,addr szCmdRunParameters,addr szCmdRunProgram,NULL
        Invoke ProccessParameters,addr szSolved1,addr szSolved2,addr szSolved3,dTerm,dUnhide,dShow,dTimeout
        ;
        ;**********************************************
        @@:
        Invoke lstrlen,addr szCopySource
        .if eax==0
            jmp @F
        .endif
        .if dCopyAfterExecute
            .if dAskCopyAfterExecute
                Invoke lstrcpy,addr szMessageBoxTitle,addr szAskForCopyTitleMain
                ;
                Invoke lstrcpy,addr szMessageBoxText,addr szAskForCopyText1
                Invoke SolveMacros,addr szCopySource,addr szSolved1 ;source, target
                Invoke lstrcat,addr szMessageBoxText,addr szSolved1 ;szCopySource
                Invoke lstrcat,addr szMessageBoxText,addr szAskForCopyText2
                Invoke SolveMacros,addr szCopyTarget,addr szSolved2
                Invoke lstrcat,addr szMessageBoxText,addr szSolved2 ;szCopyTarget
                ;
                Invoke Sleep,50
                Invoke ForceWindowToTop,hDesktopWin
                ;
                .if dGuiOnCancelCopyAfter
                    mov ecx,MB_YESNOCANCEL+MB_ICONQUESTION+MB_DEFBUTTON1+MB_TOPMOST+MB_SYSTEMMODAL
                .else
                    mov ecx,MB_YESNO+MB_ICONQUESTION+MB_DEFBUTTON1+MB_TOPMOST+MB_SYSTEMMODAL
                .endif
                Invoke MessageBox,hDesktopWin,addr szMessageBoxText,addr szMessageBoxTitle,ecx
                .if eax == IDNO
                    jmp @F
                .elseif eax == IDCANCEL
                    mov eax,TRUE
                    jmp ExitAndCallMain
                .endif
            .endif
            Invoke SolveMacros,addr szCopySource,addr szSolved1 ;source, target
            Invoke SolveMacros,addr szCopyTarget,addr szSolved2
            Invoke CopyFiles,addr szSolved1,addr szSolved2,dCopySubFolders,dCopyOverwrite
        .endif
        @@:
    .endif
    ;
ExitRunCmd:
    xor eax,eax
ExitAndCallMain:
    ret
RunCmdAtMainEntry               endp
; #########################################################################
                    ;GetCommandLine
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
CheckCommandLineArgs            proc
                                ; Return TRUE on Success
                                ;
                                LOCAL     dTotalParameters:DWORD
                                LOCAL     dShellExecuteParameters:DWORD
        ;
        mov dShellExecuteParameters,2
        Invoke FillBuffer,addr szCommandLineParameters,sizeof szCommandLineParameters,0
        Invoke GetCmdLine,1,addr szCmdRunProgram
        .if eax!=1 ;1=Sucess,2=No Argument,3=NoMatch
            xor eax,eax ;No RunProgram parameter = Help Parameter
            jmp Exit_CheckCommandLineArgs_Error
        .else
            lea ecx,szCmdRunProgram
            mov dl, byte ptr[ecx] ;"RunHide.exe ?"
            add ecx,1
            mov dh, byte ptr[ecx] ;"RunHide.exe /?"
            .if ( (dl=='?') || (dh=='?') )
                xor eax,eax ;Help Parameter
            .else
                ;Find the Total Number of Parameters. Works backwards seaching for the RunHide Operational Parameters. (max 4: TERM, UNHIDE, TIMEOUT, and SHOW)
                mov dTotalParameters,1
                @@:
                    Invoke GetCmdLine,dTotalParameters,addr szCmdRunParameters
                    .if eax!=2 ;1=Sucess,2=No Argument,3=NoMatch
                        add dTotalParameters,1
                        jmp @B
                    .else
                        sub dTotalParameters,1
                    .endif
                .if dTotalParameters==1
                    jmp Exit_CheckCommandLineArgs
                .endif
                ;Here we are in with the Total Number of parameters.
                ;Let's Check Trim the Last 4 of them to find out the Operational Parameters.
                TrimOperationalParameters:
                ;
                Invoke GetCmdLine,dTotalParameters,addr szCmdRunParameters
                .if eax==2;1=Sucess,2=No Argument,3=NoMatch
                    jmp Exit_CheckCommandLineArgs
                .else
                    ;
                    lea eax,szCmdRunParameters
                    mov pParameters,eax
                    ;
                    ;Invoke lstrcpyn, addr szCmdCompare,addr szCmdRunParameters, sizeof szCmdTerm
                    Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szCmdRunParameters, -1,addr szCmdTerm, -1
                    .if eax==2 ;2   The string pointed to by lpString1 is equal in lexical value to the string pointed to by lpString2.
                        mov dTerm,eax
                        ;mov eax,pParameters
                        ;mov byte ptr[eax],0
                    .else
                        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szCmdRunParameters, -1,addr szCmdUnhide, -1
                        .if eax==2
                            mov dUnhide,eax
                            mov dTerm,eax
                            ;Invoke MessageBox,NULL,addr szCmdRunParameters,addr szCmdUnhide,NULL
                            ;mov eax,pParameters
                            ;mov byte ptr[eax],0
                        .else
                            Invoke lstrcpyn,addr szCmdCompare,pParameters,sizeof szCmdTimeout
                            ;
                            Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szCmdCompare, -1,addr szCmdTimeout, -1
                            .if eax==2
                                ;
                                ;Detect first Numeric Character
                                xor edx,edx
                                mov ecx,pParameters
                                add ecx,sizeof szCmdTimeout
                                @@:
                                    mov al, byte ptr[ecx]
                                    .if ( al==0 )
                                        xor ecx,ecx ;Nothing to do.
                                        jmp @F
                                    .elseif( (al>='0') && (al<='9') )
                                        jmp @F
                                    .else
                                        add ecx,1
                                        jmp @B
                                    .endif
                                @@:
                                .if ecx==0
                                    xor eax,eax
                                .else
                                    ;Detect Non-numeric char
                                    mov edx,ecx
                                    @@:
                                        mov al, byte ptr[edx]
                                        .if( (al>='0') && (al<='9') )
                                            add edx,1
                                            jmp @B
                                        .else
                                            mov byte ptr[edx],0
                                            jmp @F
                                        .endif
                                    @@:
                                    Invoke AscciiToDword,ecx
                                .endif
                                ;
                                mov dTimeout,eax
                                ;mov eax,pParameters
                                ;mov byte ptr[eax],0
                                ;
                            .else
                                ;
                                Invoke lstrcpyn,addr szCmdCompare,pParameters,sizeof szCmdShow
                                Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szCmdCompare, -1,addr szCmdShow, -1
                                .if eax!=2
                                    ;No More Operational Parameters.
                                    ;add dTotalParameters,1
                                    jmp Copy_CmdParameters
                                .else
                                    ;
                                    ;Detect first Numeric Character or T char
                                    xor eax,eax
                                    xor edx,edx
                                    mov ecx,pParameters
                                    add ecx,sizeof szCmdShow
                                    @@:
                                        mov al, byte ptr[ecx]
                                        .if ( al==0 )
                                            xor ecx,ecx ;Nothing to do.
                                            jmp @F
                                        .elseif( (al=='t') ||  (al=='T') )
                                            mov al,1
                                            jmp @F
                                        .elseif( (al>='0') && (al<='9') )
                                            sub al,'0'
                                            jmp @F
                                        .else
                                            add ecx,1
                                            jmp @B
                                        .endif
                                    @@:
                                    mov dShow,eax
                                    ;mov eax,pParameters
                                    ;mov byte ptr[eax],0
                                .endif
                            .endif
                        .endif
                    .endif ;szCmdTerm
                .endif ;Exist dReverseOrderParameter

                sub dTotalParameters,1
                .if dTotalParameters == 0
                    jmp Exit_CheckCommandLineArgs
                .else
                    jmp TrimOperationalParameters
                .endif
                ;
            .endif
        .endif
        Copy_CmdParameters:
            mov eax,dTotalParameters
            .if dShellExecuteParameters > eax
                jmp Exit_CheckCommandLineArgs
            .endif

            Invoke GetCmdLine,dShellExecuteParameters,addr szCmdRunParameters
            ;Invoke MessageBox,NULL,addr szCommandLineParameters,addr szCmdRunParameters,NULL
            Invoke lstrlen,addr szCommandLineParameters
            .if eax!=0
                lea ecx,szCommandLineParameters
                add ecx,eax
                mov byte ptr [ecx],' '
                inc ecx
                mov byte ptr [ecx],00h
            .endif
            Invoke lstrcat,addr szCommandLineParameters,addr szCmdRunParameters
            add dShellExecuteParameters,1
            ;Invoke MessageBox,NULL,addr szCommandLineParameters,addr szCmdRunProgram,NULL
            jmp Copy_CmdParameters
            ;Invoke MessageBox,NULL,addr szCommandLineParameters,addr szCmdRunProgram,NULL
        Exit_CheckCommandLineArgs:
            xor eax,eax
            add eax,1
        Exit_CheckCommandLineArgs_Error:
        ;
    ret
CheckCommandLineArgs            endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;GetCommandLine  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
GetCmdLine                      proc dArgNumber:DWORD,pCstring:DWORD ;1=Sucess,2=No Argument,3=NoMatch
                                ;
                                ; 1 = successful operation
                                ; 2 = no argument exists at specified arg number
                                ; 3 = non matching quotation marks
        ;
        add dArgNumber, 1
        Invoke GetCommandLine
        Invoke GetArgNumber,eax,pCstring,dArgNumber,0
        ;
        .if eax >= 0
            mov ecx, pCstring
            .if byte ptr [ecx] != 0
                mov eax,1       ; successful operation
            .else
                mov eax,2       ; no argument at specified number
            .endif
        .elseif eax == -1
            mov eax, 3          ; non matching quotation marks
        .endif
        ;
    ret
GetCmdLine                      endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;GetArgNumber  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
GetArgNumber                    proc src:DWORD,dst:DWORD,num:DWORD,npos:DWORD
                                ;
    ;
    comment * ------------------------------------------
        Return values in EAX
        --------------------
        >0 = there is a higher number argument available
        0  = end of command line has been found
        -1 = a non matching quotation error occurred

        conditions of 0 or greater can have argument
        content which can be tested as follows.

        Test Argument for content
        -------------------------
        If the argument is empty, the first BYTE in the
        destination buffer is set to zero

        mov eax, destbuffer     ; load destination buffer
        cmp BYTE PTR [eax], 0   ; test its first BYTE
        je no_arg               ; branch to no arg handler
        print destbuffer,13,10  ; display the argument

        NOTE : A pair of empty quotes "" returns ascii 0
               in the destination buffer
        ----------------------------------------------- *
        ;
        push ebx
        push esi
        push edi
        ;
        mov esi, src
        add esi, npos
        mov edi, dst
        ;
        mov BYTE PTR [edi], 0           ; set destination buffer to zero length
        xor ebx, ebx
        ;
        ; *********
        ; scan args
        ; *********
        bcscan:
            movzx eax, BYTE PTR [esi]
            add esi, 1
            cmp BYTE PTR [abntbl+eax], 1    ; delimiting character
            je bcscan
            cmp BYTE PTR [abntbl+eax], 2    ; ASCII zero
            je quit
            ;
            sub esi, 1
            add ebx, 1
            cmp ebx, num                    ; copy next argument if number matches
            je cparg
        ;
        gcscan:
            movzx eax, BYTE PTR [esi]
            add esi, 1
            cmp BYTE PTR [abntbl+eax], 0    ; OK character
            je gcscan
            cmp BYTE PTR [abntbl+eax], 2    ; ASCII zero
            je quit
            cmp BYTE PTR [abntbl+eax], 3    ; quotation
            je dblquote
            jmp bcscan                      ; return to delimiters
        ;
        dblquote:
            add esi, 1
            cmp BYTE PTR [esi], 0
            je qterror
            cmp BYTE PTR [esi], 34
            jne dblquote
            add esi, 1
            jmp bcscan                      ; return to delimiters
        ;
        ; ********
        ; copy arg
        ; ********
        cparg:
            xor eax, eax
            xor edx, edx
            cmp BYTE PTR [esi+edx], 34      ; if 1st char is a quote
            je cpquote                      ; jump to quote copy
            @@:
                mov al, [esi+edx]               ; copy normal argument to buffer
                mov [edi+edx], al
                test eax, eax
                jz quit
                add edx, 1
                cmp BYTE PTR [abntbl+eax], 1
            jl @B
            mov BYTE PTR [edi+edx-1], 0     ; append terminator
            jmp next_read
        ;
        ; ********************
        ; copy quoted argument
        ; ********************
        cpquote:
            add esi, 1
            @@:
                mov al, [esi+edx]               ; strip quotes and copy contents to buffer
                test al, al
                jz qterror
                cmp al, 34
                je write_zero
                mov [edi+edx], al
                add edx, 1
                jmp @B
        ;
        write_zero:
            add edx, 1                      ; correct EDX for final exit position
            mov BYTE PTR [edi+edx-1], 0     ; append terminator
        ;
        jmp next_read                    ; jump to next read setup
        ;
        ; *****************
        ; set return values
        ; *****************
        qterror:
            mov eax, -1                     ; quotation error
            jmp rstack
        ;
        quit:
            xor eax, eax                    ; set EAX to zero for end of source buffer
            jmp rstack
        ;
        next_read:
            mov eax, esi
            add eax, edx
            sub eax, src
        ;
        rstack:
            pop edi
            pop esi
            pop ebx
        ;
    ret
GetArgNumber                    endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;Process Parameters
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
ProccessParameters              proc pszApp:DWORD,pszPara:DWORD,pszWrkDir:DWORD,dParaTerm:DWORD,dParaUnhide:DWORD,dParaDebugShow:DWORD,dParaTimeout:DWORD
                                LOCAL d1000Times:DWORD
                                LOCAL dShowWindow:DWORD
                                LOCAL dwTimeOut:DWORD
                                LOCAL dForceWait:DWORD
                                LOCAL dProcessID:DWORD
        ;
            ;Invoke DwordToAsccii,dParaTimeout,addr szOutputDebugString
            ;Invoke OutputDebugString,addr szOutputDebugString
        ;
        mov eax,dParaTimeout
        .if eax!=0 && eax!=-1 ;Timeout is not 0 or -1
            xor edx,edx
            mov ecx,1000  ;second to miliseconds
            mul ecx
        .endif
        mov dwTimeOut,eax
        ;
        xor ecx,ecx
        .if dParaUnhide
            mov cx,SW_SHOWNORMAL
        .else
            mov cx,SW_HIDE
        .endif
        mov dShowWindow,ecx
        ;
        xor eax,eax
        mov dProcessID,eax
        ;
        .if dParaDebugShow==NULL && dParaUnhide==NULL
            .if dParaTerm
                mov d1000Times,1000 ;Let's asume there are less than 1000 applications Open.
                @@:
                    Invoke FindApplicationProcessID,pszApp,2,NULL ;dWhatToDoIfFound: Hide=0, Unhide=1, Kill=2
                    .if eax!=NULL
                        sub d1000Times,1
                        jnz @B
                    .endif
                jmp Exit_ProccessParameters
             .endif
        .endif
        ;
        .if dParaUnhide || dParaDebugShow
            Invoke FindApplicationProcessID,pszApp,1,NULL     ;dWhatToDoIfFound: Hide=0, Unhide=1, Kill=2
            .if eax ;Running process found
                mov dProcessID,eax
                jmp Exit_ProccessParameters   ;Already running? Just update ProcessID and return.
            .else
                .if dParaUnhide!=0 && dParaDebugShow == 0
                    jmp Exit_ProccessParameters   ;Trying to unhide a process that does not exist.
                .endif
            .endif
        .endif
        ;Invoke OutputDebugString,pszApp
        .if dProcessID==NULL
            ;Invoke DwordToAsccii,ecx,addr szOutputDebugString
            ;Invoke OutputDebugString,addr szOutputDebugString
            ;Invoke lstrlen,pszPara
            ;mov ecx,eax
            ;mov ecx,dShowWindow
            ;add ecx,dParaDebugShow
            ;Invoke DwordToAsccii,ecx,addr szDebugBuffer
            ;Invoke MessageBox,NULL,addr szDebugBuffer,pszApp,NULL
            ;Invoke Sleep,1000
            Invoke ShellExecuteWait,pszApp,pszPara,pszWrkDir,dShowWindow,dwTimeOut,dParaDebugShow
            mov dProcessID,eax
        .endif;
        ;
        Exit_ProccessParameters:
        mov eax,dProcessID
    ret
ProccessParameters              endp
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;ShellExecuteWait  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
ShellExecuteWait                proc pszApplicationPath:DWORD,pszParameters:DWORD,pszWorkingDir:DWORD,dShowWindow:DWORD,dwTimeOut:DWORD,dParaDebugShow:DWORD
                                LOCAL   dIsGUIApp:DWORD
                                LOCAL   Sh_st_info:STARTUPINFO
                                LOCAL   Sh_pr_info:PROCESS_INFORMATION
                                LOCAL   dCreationFlags:DWORD
                                ;
                                LOCAL   Sh_sfi:SHFILEINFO
                                LOCAL   dReturnProcessID:DWORD
                                ;
                                LOCAL   szTexto[1024]:BYTE
                                ;LOCAL  szFullPathLocal[4096]:BYTE
                                ;LOCAL  szCurrentDirectory[4096]:BYTE
                                LOCAL   pFullPathLocal:DWORD
                                ;LOCAL  szExtentionDotExe[5]:BYTE
                                ;szText  szExtentionDotExe,".exe"
;----------------------------------------------------------------------------;
; This is a translation of a shell script originally written by Edgar Hansen ;
; in GoASM code. Translation by Paul E. Brennick  <pebrennick@verizon.net>   ;
;                                                                            ;
; Shell:      This function will run an executable and wait until it has     ;
;             finished before continuing. The function provides for a time   ;
;             out for a maximum wait period.                                 ;
; Parameters: lpfilename = The fully qualified path to an executable to run  ;
;             dwTimeOut = The amount of time in milliseconds to wait, -1 for ;
;             no timeout                                                     ;
; Returns:    0 if successful, STATUS_TIMEOUT if the timeout has elapsed     ;
;             -1 if there was an error creating the process                  ;
;                                                                            ;
;----------------------------------------------------------------------------;
        ;
        Invoke FillBuffer,addr szFullPathApp,sizeof szFullPathApp,0
        Invoke FillBuffer,addr szFullPathLocal,sizeof szFullPathLocal,0
        Invoke FillBuffer,addr szFullPathWorkingDir,sizeof szFullPathWorkingDir,0
        xor eax,eax
        mov dReturnProcessID,eax
        Invoke lstrlen,pszApplicationPath
        .if eax==0
            jmp ExitShellExecuteWait
        .endif
        ;
        Invoke SearchPath,NULL,pszApplicationPath,addr szExtentionDotExe,sizeof szFullPathLocal,addr szFullPathLocal,addr pFullPathLocal
        .if eax!=0 ;If the function fails, the return value is zero. To get extended error information, call GetLastError.
            ;
            xor eax,eax
            mov dIsGUIApp,eax
            mov eax,pFullPathLocal
            lea ecx,szFullPathLocal
            sub eax,ecx
            Invoke lstrcpyn,addr szFullPathWorkingDir,addr szFullPathLocal,eax ;Target,Source,BytesToCopy
            .if pszWorkingDir != NULL
                Invoke lstrlen,pszWorkingDir
                .if eax!=0
                    Invoke lstrcpy,addr szFullPathWorkingDir,pszWorkingDir ;Target,Source,BytesToCopy
                .endif
            .endif
            ;
            Invoke SHGetFileInfo,addr szFullPathLocal,0,addr Sh_sfi,sizeof Sh_sfi,SHGFI_EXETYPE
            .if eax==0
                mov dCreationFlags,eax ;Nonexecutable file or an error condition.
            .else
                mov ecx,eax
                shr ecx,16
                .if (ax == 454EH) ;Windows application: LOWORD = NE (4E45H) or PE (5045H),HIWORD = 3.0, 3.5, or 4.0
                   mov dIsGUIApp,eax
                    ;WIN16_GUI = NE Win16  454EH
                   mov dCreationFlags,NORMAL_PRIORITY_CLASS+CREATE_DEFAULT_ERROR_MODE+CREATE_SEPARATE_WOW_VDM ;+EXTENDED_STARTUPINFO_PRESENT
                .elseif(ax==5A4DH)  ;MS-DOS .EXE, .COM or .BAT file: LOWORD = MZ (4D5AH),HIWORD = 0
                    ;WIN16_DOS = MZ MS-DOS 5A4DH
                    mov dCreationFlags,NORMAL_PRIORITY_CLASS+CREATE_DEFAULT_ERROR_MODE+CREATE_SEPARATE_WOW_VDM+CREATE_NEW_CONSOLE ;+EXTENDED_STARTUPINFO_PRESENT
                .elseif( (ax==4550H) && (cx==0) ) ;Win32 console application: LOWORD = PE(5045H),HIWORD = 0
                    ;WIN32_CON = PE Win32 CONSOLE 4550H
                    mov dCreationFlags,NORMAL_PRIORITY_CLASS+CREATE_NEW_CONSOLE ;+EXTENDED_STARTUPINFO_PRESENT
                .else
                    mov dIsGUIApp,eax
                    ;WIN32_GUI = PE Win32  4550H
                    mov dCreationFlags,NORMAL_PRIORITY_CLASS+CREATE_DEFAULT_ERROR_MODE ;+EXTENDED_STARTUPINFO_PRESENT
                .endif
            .endif
            ;
        .else
            ;
            Invoke GetLastError
            mov ecx,eax
            Invoke FormatMessage,FORMAT_MESSAGE_FROM_SYSTEM, 0,ecx, 0,addr szFullPathLocal,MAX_PATH, 0
            .if eax!=0
                lea ecx,szFullPathLocal
                add ecx,eax
                mov byte ptr[ecx],10
                inc eax
                mov byte ptr[ecx],0
                Invoke  lstrcat,addr szFullPathLocal,pszApplicationPath
                Invoke MessageBox,NULL,addr szFullPathLocal,addr szErrorExecuteTitle,MB_OK+MB_ICONHAND ;+MB_SYSTEMMODAL
            .endif
            Invoke lstrcpy,addr szFullPathLocal,pszApplicationPath
            ;
        .endif
        ;
        Invoke FillBuffer,addr Sh_st_info,sizeof STARTUPINFO,NULL
        mov     dword ptr[Sh_st_info.cb],sizeof STARTUPINFO
        Invoke  GetStartupInfoA, addr Sh_st_info
        ;Invoke MessageBox,NULL,Sh_st_info.lpDesktop,Sh_st_info.lpTitle,NULL
        mov     dword ptr[Sh_st_info.lpReserved],NULL
        .if dIsGUIApp!=0
            mov dword ptr[Sh_st_info.lpTitle],NULL  ;This parameter must be NULL for GUI
        .else
            mov eax,pszApplicationPath
            mov dword ptr[Sh_st_info.lpTitle],eax  ;This parameter must be NULL for GUI
            ;Invoke MessageBox,NULL,Sh_st_info.lpDesktop,Sh_st_info.lpTitle,NULL
        .endif
        ;
        mov     dword ptr[Sh_st_info.lpReserved],0
        mov     dword ptr[Sh_st_info.dwFlags],STARTF_USESHOWWINDOW+STARTF_FORCEOFFFEEDBACK   ;STARTF_USESHOWWINDOW ;EXTENDED_STARTUPINFO_PRESENT
        mov     eax,dShowWindow
        mov     word ptr[Sh_st_info.wShowWindow],ax ;SW_MINIMIZE ;SW_MINIMIZE ; SW_MAXIMIZE ;SW_HIDE=0 ;SW_SHOW=5
        mov     dword ptr[Sh_st_info.cbReserved2],NULL
        mov     dword ptr[Sh_st_info.lpReserved2],NULL
        ;
        Invoke lstrcpy,addr szFullPathApp,addr szFullPathLocal ;Target,Source
        ;
        xor edx,edx
        mov eax,pszParameters
        mov al, byte ptr[eax]
        .if (al!=0)
            Invoke lstrlen,addr szFullPathLocal
            lea ecx,szFullPathLocal
            mov byte ptr[eax+ecx],' '
            mov byte ptr[eax+ecx+1],0
            Invoke lstrcat,addr szFullPathLocal,pszParameters
        .endif
        ;
        ;
        mov eax,dIsGUIApp
        .if (ax == 454EH)   ;Windows application: LOWORD = NE (4E45H) or PE (5045H),HIWORD = 3.0, 3.5, or 4.0
            xor ecx,ecx ;If the executable module is a 16-bit application, lpApplicationName should be NULL
        .else
            lea ecx,szFullPathApp
        .endif
        ;
        ;
        invoke  CreateProcessA,ecx,addr szFullPathLocal,0,0,0,dCreationFlags,0,addr szFullPathWorkingDir, \
                addr Sh_st_info, addr Sh_pr_info ;Handles in STARTUPINFO or STARTUPINFOEX must be closed with CloseHandle when they are no longer needed.
        .if eax!=0 ;If the function succeeds, the return value is nonzero.
            mov eax,Sh_pr_info.dwProcessId ;Sh_pr_info.hProcess
            mov dReturnProcessID,eax
            ;
            .if dParaDebugShow==NULL && dShowWindow==SW_HIDE
                invoke WaitForInputIdle,Sh_pr_info.hProcess,INFINITE
                mov eax,dShowWindow
                .if ax==SW_HIDE
                    ;TODO: Better way of ensure CreateProccessA run exe in hidden mode?
                    Invoke FindApplicationProcessID,pszApplicationPath,0,NULL  ;dWhatToDoIfFound: Hide=0, Unhide=1, Kill=2
                .endif
            .endif
            .if dwTimeOut == 0 || dwTimeOut == -1
                jmp CleanShellExecuteWait
            .else
                mov ecx,dwTimeOut
                Invoke WaitForSingleObject,[Sh_pr_info.hProcess],ecx
                .if eax==WAIT_TIMEOUT
                    ;Invoke DwordToAsccii,ecx,addr szTexto
                    ;Invoke MessageBox,NULL,addr szTexto,addr szTexto,NULL
                    .if dParaDebugShow==NULL
                        Invoke TerminateProcess,[Sh_pr_info.hProcess],0
                        xor eax,eax
                        mov dReturnProcessID,eax
                    .endif
                .endif
                jmp CleanShellExecuteWait
            .endif
            @@:
            CleanShellExecuteWait:
            push eax
            Invoke  CloseHandle, [Sh_pr_info.hThread]
            Invoke  CloseHandle, [Sh_pr_info.hProcess]
            pop eax
        .endif
        ;
        ExitShellExecuteWait:
        mov eax,dReturnProcessID
    ret
ShellExecuteWait                endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;FindApplicationProcessID  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
FindApplicationProcessID        proc pszApplicationPath:DWORD,dWhatToDoIfFound:DWORD,dOutWinHandle:DWORD ;dWhatToDoIfFound: Update OutWinHandle=-1, Hide=0, Unhide=1, Kill=2
                                LOCAL hWin:DWORD
                                LOCAL bLoop:BOOL
                                LOCAL dThreadID:HANDLE
                                LOCAL dReturnProccessID:HANDLE
                                LOCAL pe32:PROCESSENTRY32
                                LOCAL hProcessesSnap:HANDLE
                                ;
                                LOCAL Sh_st_info :STARTUPINFO
                                LOCAL Sh_sfi     :SHFILEINFO
                                ;
                                ;szText  szExtentionDotExe,".exe"

                                LOCAL pFullPathLocal:DWORD
                                LOCAL dlParam[3]:DWORD
        ;
        Invoke FillBuffer,addr szFullPathLocal,sizeof szFullPathLocal,0
        mov dReturnProccessID,NULL
        mov pFullPathLocal,0
        ;
        Invoke SearchPath,NULL,pszApplicationPath,addr szExtentionDotExe,sizeof szFullPathLocal,addr szFullPathLocal,addr pFullPathLocal
        ;
        ;
        ;You should call this function from a background thread. Failure to do so could cause the UI to stop responding.
        Invoke  SHGetFileInfo,addr szFullPathLocal,0,addr Sh_sfi,sizeof Sh_sfi,SHGFI_EXETYPE
                ;Return:
                ;0                                              [Nonexecutable file or an error condition]
                ;LOWORD = NE or PE and HIWORD = Windows version [Windows application]
                ;LOWORD = MZ and HIWORD = 0                     [MS-DOS .exe or .com file]
                ;LOWORD = PE and HIWORD = 0                     [Console application or .bat file]
        .if eax==0
            jmp ExitFindApplicationProcessID
        .endif
        mov ecx,eax
        shr ecx,16
        .if cx==0
            ;ax==4550H -> WIN32_CON = PE [Console application or .bat file]: 4550H from 45H=69dec, 50H=80dec, 69dec ASCII = E and 80dec ASCII = P
            ;ax==5A4DH ->                [MS-DOS .exe or .com file]:                                          90dec ASCII = Z and 77dec ASCII = M
            mov     dword ptr [Sh_st_info.cb],sizeof STARTUPINFO
            Invoke  GetStartupInfoA, addr Sh_st_info
            mov ecx,dword ptr [Sh_st_info.lpTitle]
            mov pszApplicationPath,ecx
            Invoke FindWindow,NULL,pszApplicationPath
            .if eax!=0
                mov hWin,eax
                Invoke GetWindowThreadProcessId,hWin,addr dReturnProccessID
                .if dWhatToDoIfFound==-1
                    mov edx,dOutWinHandle
                    mov DWORD PTR edx,hWin
                .elseif dWhatToDoIfFound==0 ;dWhatToDoIfFound Hide=0 Unhide=1 Kill=2
                    Invoke ForceWindowToTop,hWin
                    Invoke ShowWindow,hWin,SW_HIDE ;SW_SHOW
                    ;mov dReturnProccessID,eax
                .elseif dWhatToDoIfFound==1
                    Invoke ShowWindow,hWin,SW_SHOW
                    Invoke ForceWindowToTop,hWin
                    ;mov dReturnProccessID,eax
                .else
                    ;Invoke GetWindowThreadProcessId,hWin,addr dReturnProccessID
                    Invoke OpenProcess,PROCESS_ALL_ACCESS,FALSE,dReturnProccessID
                    .if eax!=NULL
                        push eax
                        Invoke TerminateProcess,eax, 0
                        pop eax
                        Invoke CloseHandle,eax
                        ;mov dReturnProccessID,
                    .endif
                .endif
            .endif
            ;
        .else ;.if cx==0
            ;cx=Window Version, ax = NE or PE       [Windows application: WIN32_EXE]
            mov dReturnProccessID,NULL
            ;
            Invoke CreateToolhelp32Snapshot,TH32CS_SNAPPROCESS,0 ;Takes a snapshot of the specified processes, as well as the heaps, modules, and threads used by these processes.
            .if eax == INVALID_HANDLE_VALUE
                xor eax,eax
                jmp ExitFindApplicationProcessID
            .endif
            mov hProcessesSnap,eax
            mov pe32.dwSize,SIZEOF PROCESSENTRY32
            mov bLoop,TRUE
            ;
            Invoke Process32First,hProcessesSnap,addr pe32
            .if eax==FALSE ;Returns TRUE if the first entry of the process list has been copied to the buffer or FALSE otherwise.
                Invoke CloseHandle,hProcessesSnap
                xor eax,eax
                jmp ExitFindApplicationProcessID
            .else
                ;=========================================
                ;TODO:
                ;pe32.szExeFile[MAX_PATH] The name of the executable file for the process.
                ; To retrieve the full path to the executable file, call the Module32First function
                ; and check the szExePath member of the MODULEENTRY32 structure that is returned.
                ; However, if the calling process is a 32-bit process, you must call the QueryFullProcessImageName function
                ; to retrieve the full path of the executable file for a 64-bit process.
                ;=========================================
                .while bLoop
                    Invoke CompareString,LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr pe32.szExeFile, -1, pFullPathLocal, -1
                    .if eax==2
                        ;
                        mov eax,pe32.th32ProcessID
                        mov dReturnProccessID,eax
                        ;
                        .if dWhatToDoIfFound == 2 ;dWhatToDoIfFound Hide=0 Unhide=1 Kill=2
                            Invoke OpenProcess,PROCESS_TERMINATE,FALSE,dReturnProccessID ;PROCESS_TERMINATE, PROCESS_ALL_ACCESS
                            .if eax==NULL
                                Invoke OpenProcess,PROCESS_ALL_ACCESS,FALSE,dReturnProccessID ;PROCESS_TERMINATE, PROCESS_ALL_ACCESS
                            .endif
                            .if eax==NULL
                                mov dReturnProccessID,eax
                            .else
                                push eax
                                Invoke TerminateProcess,eax,0
                                .if eax==NULL ;If the function succeeds, the return value is nonzero.
                                    mov dReturnProccessID,eax
                                .endif
                                pop eax
                                Invoke CloseHandle,eax
                                ;mov dReturnProccessID,0
                            .endif
                        .else
                            ;Here we are in the Application Process
                            Invoke FindTopWindowFromProcessID,dReturnProccessID,dWhatToDoIfFound
                        .endif
                        .break
                    .endif
                    Invoke Process32Next, hProcessesSnap,addr pe32
                    mov bLoop,eax
                .endw
                Invoke CloseHandle,hProcessesSnap
            .endif
        .endif ;.if ( (ax==4550H) && (cx==0) )
        ;
        mov eax,dReturnProccessID
        ;
    ExitFindApplicationProcessID:
        ;
    ret
FindApplicationProcessID        endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
;                       FindTopWindowFromProcessID
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
FindTopWindowFromProcessID      proc dProcessId:DWORD,dWhatToDoIfFound:DWORD
                                    LOCAL hProcess:HANDLE
                                    LOCAL bLoop:BOOL
                                    LOCAL dOutWinHandle:DWORD
                                    LOCAL dlParam[3]:DWORD
        ;
        mov dOutWinHandle,NULL
        Invoke OpenProcess,PROCESS_ALL_ACCESS,FALSE,dProcessId
        mov hProcess,eax
        .if hProcess!=NULL ;If the function succeeds, the return value is an open handle to the specified process.;
            lea edx,dlParam ;dlParam[3]:DWORD
            mov eax,dProcessId
            mov dword ptr[edx],eax
            mov eax,dWhatToDoIfFound ;dWhatToDoIfFound!=2 ;dWhatToDoIfFound Hide=0 Unhide=1 Kill=2
            mov dword ptr[edx+4],eax
            lea eax,dOutWinHandle
            mov dword ptr[edx+8],eax
            .while bLoop
                ;Enumerates all top-level windows on the screen by passing the handle to each window, in turn,
                ; to an application-defined callback function. EnumWindows continues until
                ; the last top-level window is enumerated or the callback function returns FALSE.
                Invoke EnumWindows,addr EnumWindowsProc,addr dlParam ;pe32.th32ProcessID
                mov bLoop,eax ;Return value Type: BOOL If the function succeeds, the return value is nonzero.
                ;Remarks: The EnumWindows function does not enumerate child windows,
                ; with the exception of a few top-level windows owned by the system that have the WS_CHILD style.
                ;This function is more reliable than calling the GetWindow function in a loop.
                ; An application that calls GetWindow to perform this task risks being caught in an infinite loop or referencing a handle to a window that has been destroyed.
                ;Note  For Windows 8 and later, EnumWindows enumerates only top-level windows of desktop apps.
            .endw
            Invoke CloseHandle,hProcess
        .endif
    mov eax,dOutWinHandle
    ret
FindTopWindowFromProcessID      endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;EnumWindowsProc  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
EnumWindowsProc                 proc hWinInTurn:DWORD,lparam:DWORD
                                    ;
                                    LOCAL dWinMainThreadID:DWORD
                                    LOCAL dWinInTurnThreadID:DWORD
                                    LOCAL dWinInTurnProcessID:DWORD
                                    LOCAL dIsChildWindow:DWORD
                                    ;
                                    LOCAL dThreadID:DWORD
                                    LOCAL dProcessID:DWORD
                                    LOCAL dWhatToDoIfFound:DWORD ;dWhatToDoIfFound Hide=0 Unhide=1 Kill=2
                                    LOCAL dOutWinHandle:DWORD
        ;
        mov edx,lparam
        ;mov eax,dword ptr[edx]
        ;mov dThreadID,eax
        mov eax,dword ptr[edx]
        mov dProcessID,eax
        mov eax,dword ptr[edx+4]
        mov dWhatToDoIfFound,eax
        mov eax,dword ptr[edx+8]
        mov dOutWinHandle,eax
        ;
        Invoke GetProcessMainThreadID,dProcessID
        mov dWinMainThreadID,eax
        ;
        mov dIsChildWindow,FALSE
        Invoke GetWindowThreadProcessId,hWinInTurn,addr dWinInTurnProcessID
        .if eax==NULL
            jmp ExitEnumWindowsProc
        .endif
        mov dWinInTurnThreadID,eax
        mov edx,dProcessID
        .if (dWinInTurnProcessID!=edx)
            ;Filter out every other window where their PID is not dProcessID
            mov dIsChildWindow,TRUE
        .else
            .if dWinMainThreadID
                mov edx,dWinInTurnThreadID
                .if dWinMainThreadID!=edx
                    mov dIsChildWindow,TRUE
                ;.else
                    ;Invoke DwordToAsccii,dWinMainThreadID,addr szOutputDebugString
                    ;Invoke OutputDebugString,addr szOutputDebugString
                    ;Invoke GetWindowText,hWinInTurn,addr szOutputDebugString,sizeof szOutputDebugString
                    ;Invoke MessageBox,hWinInTurn,addr szOutputDebugString,addr szOutputDebugString,NULL
                    ;Invoke OutputDebugString,addr szOutputDebugString
                .endif
            .else
                Invoke GetWindow,hWinInTurn,GW_OWNER
                .if eax!=NULL
                    ;Filter out every windows that has OWNER
                    mov dIsChildWindow,TRUE
                .else
                    Invoke GetParent,hWinInTurn
                    .if eax!=NULL
                        ;Filter out every windows that has PARENT
                        mov dIsChildWindow,TRUE
                    .else
                        ;
                        ;DISCARDED->Invoke IsWindowVisible,hWinInTurn ;Filter does not works if the app is running hidden.
                        ;                  ;If the specified window, its parent window, its parent's parent window, and so forth,
                        ;                  ;have the WS_VISIBLE style, the return value is nonzero. Otherwise, the return value is zero.
                        ;
                        ;TODO: What happen when a Windows has not Title text?
                        Invoke GetWindowText,hWinInTurn,addr szOutputDebugString,sizeof szOutputDebugString
                         .if eax==NULL
                            ;Filter out every windows that has not Text in title.
                            mov dIsChildWindow,TRUE
                        ;.else
                            ;Invoke GetWindowText,hWinInTurn,addr szOutputDebugString,sizeof szOutputDebugString
                            ;Invoke MessageBox,hWinInTurn,addr szOutputDebugString,addr szOutputDebugString,NULL
                            ;Invoke OutputDebugString,addr szOutputDebugString
                        .endif
                    .endif
                .endif
            .endif
        .endif
        .if dIsChildWindow==FALSE
            .if dWhatToDoIfFound==-1 ;dWhatToDoIfFound: Hide=0, Unhide=1, Kill=2
                .if dOutWinHandle!=NULL
                    mov edx,dOutWinHandle
                    mov dword ptr edx,hWinInTurn
                .endif
            .elseif dWhatToDoIfFound==0 ;dWhatToDoIfFound: Hide=0, Unhide=1, Kill=2
                Invoke ShowWindow,hWinInTurn,SW_HIDE
            .elseif dWhatToDoIfFound==1
                Invoke ShowWindow,hWinInTurn,SW_SHOW
                Invoke IsIconic,hWinInTurn
                .if eax
                    Invoke ShowWindow,hWinInTurn,SW_RESTORE
                .endif
                Invoke ForceWindowToTop,hWinInTurn
            .endif
        .endif
        ;To continue enumeration, the callback function must return TRUE; to stop enumeration, it must return FALSE.
        ;EnumWindows continues until the last top-level window is enumerated or the callback function returns FALSE.
        mov eax,dIsChildWindow
        ExitEnumWindowsProc:
    ret
EnumWindowsProc                 endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;GetProcessMainThreadID  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
GetProcessMainThreadID          proc hProcessID:DWORD
                                LOCAL   dReturnThreadID:HANDLE
                                LOCAL   hThreadsSnap:HANDLE
                                LOCAL   th32:THREADENTRY32
                                LOCAL   bLoop:BOOL
                                LOCAL   hThread:HANDLE
                                LOCAL   afMinCreationTime:FILETIME
                                LOCAL   afCreationTime:FILETIME
                                LOCAL   afExitTime:FILETIME
                                LOCAL   afKernelTime:FILETIME
                                LOCAL   afUserTime:FILETIME
                                ;
        mov dReturnThreadID,NULL
        ;
        Invoke CreateToolhelp32Snapshot,TH32CS_SNAPTHREAD,0 ;Takes a snapshot of the specified processes, as well as the heaps, modules, and threads used by these processes.
        .if eax == INVALID_HANDLE_VALUE
            jmp ExitGetProcessMainThreadID
        .endif
        mov hThreadsSnap,eax
        ;
        mov afMinCreationTime.dwLowDateTime,0FFFFFFFFH
        mov afMinCreationTime.dwHighDateTime,0FFFFFFFFH
        mov th32.dwSize,SIZEOF THREADENTRY32
        mov bLoop,TRUE
        ;
        Invoke Thread32First,hThreadsSnap,addr th32
        .if eax==FALSE ;Returns TRUE if the first entry of the process list has been copied to the buffer or FALSE otherwise.
            Invoke CloseHandle,hThreadsSnap
        .else
            .while bLoop
                ;
                mov eax,hProcessID
                .if th32.th32OwnerProcessID == eax
                    Invoke OpenThread,THREAD_QUERY_INFORMATION,TRUE,th32.th32ThreadID
                    .if eax ;If the function fails, the return value is NULL
                        mov hThread,eax
                        ;
                        Invoke GetThreadTimes,hThread,addr afCreationTime,addr afExitTime,addr afKernelTime,addr afUserTime
                        .if eax ;If the function succeeds, the return value is nonzero.
                            .if afCreationTime.dwLowDateTime!=NULL && afCreationTime.dwHighDateTime!=NULL
                                Invoke CompareFileTime,addr afMinCreationTime,addr afCreationTime
                                ;Return
                                ;-1 = First file time is earlier than second file time.
                                ; 0 = First file time is equal to second file time.
                                ; 1 = First file time is later than second file time.
                                .if eax==1
                                    mov ecx,afCreationTime.dwLowDateTime
                                    mov afMinCreationTime.dwLowDateTime,ecx
                                    mov edx,afCreationTime.dwHighDateTime
                                    mov afMinCreationTime.dwHighDateTime,edx
                                    ;
                                    mov eax,th32.th32ThreadID
                                    mov dReturnThreadID,eax ;let it be main... :)
                                .endif
                            .endif
                        .endif
                        ;
                        Invoke CloseHandle,hThread
                    .endif
                .endif
                ;
                ;
                Invoke Thread32Next,hThreadsSnap,addr th32
                mov bLoop,eax
            .endw
            Invoke CloseHandle,hThreadsSnap
        .endif
        ;
        ExitGetProcessMainThreadID:
        mov eax,dReturnThreadID
    ret
GetProcessMainThreadID          endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
;                       ForceWindowToTop
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
ForceWindowToTop                proc hWnd:DWORD ;********* PUT This Windows ABOVE it all
                                    LOCAL dwCurrentThread :DWORD
                                    LOCAL dwFGThread :DWORD
                                    LOCAL dwProcessID :DWORD
                                    ;szText  szExtentionDotExe,".exe"
                                ;
        ;https://stackoverflow.com/questions/688337/how-do-i-force-my-app-to-come-to-the-front-and-take-focus
        ;
        xor eax,eax
        Invoke IsWindow,hWnd
        .if eax ;If the window handle does not identify an existing window, the return value is zero.
            Invoke GetCurrentThreadId
            mov dwCurrentThread,eax
            Invoke GetForegroundWindow
            Invoke GetWindowThreadProcessId,eax,NULL ;addr dwProcessID
            mov dwFGThread,eax
            Invoke AttachThreadInput,dwCurrentThread,dwFGThread,TRUE
            ;Invoke GetWindowText,hWnd,addr szBuffer,sizeof szBuffer
            ;Invoke MessageBox,hWnd,addr szBuffer,addr szExtentionDotExe,MB_YESNO+MB_ICONQUESTION
            ;Possible actions you may try to bring the window into focus.
            Invoke SetForegroundWindow,hWnd
            .if eax==0 ;If the window was not brought to the foreground, the return value is zero.
                Invoke Sleep,10
                Invoke SwitchToThisWindow,hWnd,FALSE
            .endif
            ;Invoke SetCapture,hWnd
            ;Invoke SetFocus,hWnd
            ;Invoke SetActiveWindow,hWnd
            ;Invoke EnableWindow,hWnd,TRUE)
            ;Invoke ShowWindowAsync,hWnd,SW_HIDE
            ;Invoke ShowWindowAsync,hWnd,SW_SHOWNORMAL
            ;Invoke ShowWindowAsync,hWnd,SW_SHOWNORMAL
            ;
            Invoke AttachThreadInput,dwCurrentThread,dwFGThread,FALSE
            ;
            Invoke GetWindowThreadProcessId,hWnd,addr dwProcessID
            mov eax,dwProcessID
;        .else
;            szText  szOutputDebugString99,"ForceWindowToTop->Invoke IsWindow,hWnd=FALSE"
;            Invoke OutputDebugString, addr szOutputDebugString99
;            xor eax,eax
        .endif
    ret
ForceWindowToTop                endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;WinMain For RunHide
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
main                            proc
                                LOCAL lpArgs:DWORD
                                ;
    ;Windows Vista and above use Segoe UI, Windows 2000 and XP used Tahoma and earlier Windows used MS Sans Serif as their default fonts
    ;
    Invoke GlobalAlloc,GMEM_FIXED or GMEM_ZEROINIT, 32
    mov lpArgs, eax
    .if bIsClarionDebugRunning
        push hIconClaDebug
    .else
        push hIcon
    .endif
    pop [eax]
    Dialog  "RunHide Launcher","MS Sans Serif",8, \        ; caption,font,pointsize
            WS_OVERLAPPED OR WS_POPUP OR WS_VISIBLE OR WS_CAPTION OR WS_SYSMENU OR DS_CENTER OR WS_EX_TOPMOST, \
            EQU_LAST_CONTROL, \                                ; control count
            0,0,512,362, \                    ; x y co-ordinates
            4096                                ; memory buffer size
            ;===================================== GROUP 1 =================================================================================
            DlgGroup "RunHide Previous Task",                                                               12 ,8  ,488,156,    IDC_GROUP_PRE
            DlgStatic "Copy from Source:",WS_CHILD OR WS_VISIBLE,                                           24 ,23 ,72 ,12 ,    IDC_PROMPT_PRE_SOURCE
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL,                    100,22 ,354,12 ,    IDC_EDIT_PRE_SOURCE
            DlgButton "...",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H,                                  458,22 ,14 ,12 ,    IDC_BUTTON_PRE_SOURCE_FILEDLG
            DlgButton "S",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H OR 11003h,                          476,22 ,14 ,12 ,    IDC_BUTTON_PRE_SOURCE_OVERWR
            DlgStatic "Copy to Target:",WS_CHILD OR WS_VISIBLE,                                             24 ,39 ,72 ,12 ,    IDC_PROMPT_PRE_TARGET
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL,                    100,38 ,354,12 ,    IDC_EDIT_PRE_TARGET
            DlgButton "...",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H,                                  458,38 ,14 ,12 ,    IDC_BUTTON_PRE_TARGET_FILEDLG
            DlgButton "O",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H OR 11003h,                          476,38 ,14 ,12 ,    IDC_BUTTON_PRE_TARGET_OVERWR
            DlgStatic "Execute:",WS_CHILD OR WS_VISIBLE,                                                    24 ,55 ,72 ,12 ,    IDC_PROMPT_PRE_EXECUTE
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL OR 8000H,           100,54 ,371,12 ,    IDC_EDIT_PRE_EXECUTE
            DlgButton "...",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H,                                  476,54 ,14 ,12 ,    IDC_BUTTON_PRE_EXECUTE_FILEDLG
            DlgStatic "Parameters:",WS_CHILD OR WS_VISIBLE,                                                 24 ,71 ,72 ,12 ,    IDC_PROMPT_PRE_PARAMS
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP  OR ES_AUTOHSCROLL OR 8000H,          100,70 ,371,12 ,    IDC_EDIT_PRE_PARAMS
            DlgButton "...",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H,                                  476,70 ,14 ,12 ,    IDC_BUTTON_PRE_PARAMS_FILEDLG
            DlgStatic "Working Directory:",WS_CHILD OR WS_VISIBLE,                                          24 ,87 ,72 ,12 ,    IDC_PROMPT_PRE_WRKDIR
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP  OR ES_AUTOHSCROLL OR 8000H,          100,86 ,371,12 ,    IDC_EDIT_PRE_WRKDIR
            DlgButton "...",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H,                                  476,86 ,14 ,12 ,    IDC_BUTTON_PRE_WRKDIR_FILEDLG
            ;
            DlgStatic "Options:",WS_CHILD OR WS_VISIBLE,                                                    24 ,104,72 ,12 ,    IDC_PROMPT_PRE_OPTIONS
            DlgCheck "Unhide",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,                           100,104,130,12 ,    IDC_CHECK_PRE_UNHIDE
            DlgCheck "Show",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,                             100,118,130,12 ,    IDC_CHECK_PRE_SHOW
            DlgCheck "Ask before 'Execute'",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,             100,132,130,12 ,    IDC_CHECK_PRE_ASK_BEFORE_EXEC
            DlgCheck "GUI on cancel 'Execute'",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,          100,146,130,12 ,    IDC_CHECK_PRE_GUI_CANCEL_EXEC
            ;
            DlgCheck "Term",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,                             230,104,130,12 ,    IDC_CHECK_PRE_TERM
            DlgCheck "Copy after 'Execute'",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,             230,118,130,12 ,    IDC_CHECK_PRE_COPY_AFTER
            DlgCheck "Ask before 'Copy after' Execute",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,  230,132,130,12 ,    IDC_CHECK_PRE_ASK_COPY_AFTER
            DlgCheck "GUI on cancel ask 'Copy after'",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,   230,146,130,12 ,    IDC_CHECK_PRE_GUI_COPY_AFTER
            ;
            DlgStatic "Time Out:",WS_CHILD OR WS_VISIBLE,                                                   360,106,48 ,12 ,    IDC_PROMPT_PRE_TIMEOUT
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_NUMBER OR 8000H,                424,104,48 ,12 ,    IDC_EDIT_PRE_TIMEOUT
            DlgComCtl "msctls_updown32", UDS_AUTOBUDDY OR UDS_SETBUDDYINT OR UDS_ARROWKEYS \
                                                                   OR UDS_ALIGNRIGHT OR UDS_NOTHOUSANDS,    424,104,24 ,12 ,    IDC_SPINBOX_PRE_TIMEOUT
            DlgCheck "Copy before 'Execute'",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,            360,118,130,12 ,    IDC_CHECK_PRE_COPY_BEFORE
            DlgCheck "Ask before 'Copy before' Execute",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H, 360,132,130,12 ,    IDC_CHECK_PRE_ASK_COPY_BEFORE
            DlgCheck "GUI on cancel ask 'Copy before'",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,  360,146,130,12 ,    IDC_CHECK_PRE_GUI_COPY_BEFORE
            ;
            ;===================================== GROUP 2 ==================================================================== ;
            DlgGroup "RunHide Main Task",                                                                   12 ,172,488,156,    IDC_GROUP_MAIN
            DlgStatic "Copy from Source:",WS_CHILD OR WS_VISIBLE,                                           24 ,187,72 ,12 ,    IDC_PROMPT_MAIN_SOURCE
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL,                    100,186,354,12 ,    IDC_EDIT_MAIN_SOURCE
            DlgButton "...",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H,                                  458,186,14 ,12 ,    IDC_BUTTON_MAIN_SOURCE_FILEDLG
            DlgButton "S",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H OR 11003h,                          476,186,14 ,12 ,    IDC_BUTTON_MAIN_SOURCE_OVERWR
            DlgStatic "Copy to Target:",WS_CHILD OR WS_VISIBLE,                                             24 ,203,72 ,12 ,    IDC_PROMPT_MAIN_TARGET
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL,                    100,202,354,12 ,    IDC_EDIT_MAIN_TARGET
            DlgButton "...",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H,                                  458,202,14 ,12 ,    IDC_BUTTON_MAIN_TARGET_FILEDLG
            DlgButton "O",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H OR 11003h,                          476,202,14 ,12 ,    IDC_BUTTON_MAIN_TARGET_OVERWR
            DlgStatic "Execute:",WS_CHILD OR WS_VISIBLE,                                                    24 ,219,72 ,12 ,    IDC_PROMPT_MAIN_EXECUTE
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL OR 8000H,           100,218,371,12 ,    IDC_EDIT_MAIN_EXECUTE
            DlgButton "...",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H,                                  476,218,14 ,12 ,    IDC_BUTTON_MAIN_EXECUTE_FILEDLG
            DlgStatic "Parameters:",WS_CHILD OR WS_VISIBLE,                                                 24 ,235,72 ,12 ,    IDC_PROMPT_MAIN_PARAMS
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP  OR ES_AUTOHSCROLL OR 8000H,          100,234,371,12 ,    IDC_EDIT_MAIN_PARAMS
            DlgButton "...",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H,                                  476,234,14 ,12 ,    IDC_BUTTON_MAIN_PARAMS_FILEDLG
            DlgStatic "Working Directory:",WS_CHILD OR WS_VISIBLE,                                          24 ,251,72 ,12 ,    IDC_PROMPT_MAIN_WRKDIR
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_AUTOHSCROLL OR 8000H,           100,250,371,12 ,    IDC_EDIT_MAIN_WRKDIR
            DlgButton "...",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H,                                  476,250,14 ,12 ,    IDC_BUTTON_MAIN_WRKDIR_FILEDLG
            ;
            DlgStatic "Options:",WS_CHILD OR WS_VISIBLE,                                                    24 ,268,72 ,12 ,    IDC_PROMPT_MAIN_OPTIONS
            DlgCheck "Unhide",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,                           100,268,130,12 ,    IDC_CHECK_MAIN_UNHIDE
            DlgCheck "Show",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,                             100,282,130,12 ,    IDC_CHECK_MAIN_SHOW
            DlgCheck "Ask before 'Execute'",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,             100,296,130,12 ,    IDC_CHECK_MAIN_ASK_BEFORE_EXEC
            DlgCheck "GUI on cancel 'Execute'",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,          100,310,130,12 ,    IDC_CHECK_MAIN_GUI_CANCEL_EXEC
            ;
            DlgCheck "Term",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,                             230,268,130,12 ,    IDC_CHECK_MAIN_TERM
            DlgCheck "Copy after 'Execute'",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,             230,282,130,12 ,    IDC_CHECK_MAIN_COPY_AFTER
            DlgCheck "Ask before 'Copy after' Execute",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,  230,296,130,12 ,    IDC_CHECK_MAIN_ASK_COPY_AFTER
            DlgCheck "GUI on cancel ask 'Copy after'",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,   230,310,130,12 ,    IDC_CHECK_MAIN_GUI_COPY_AFTER
            ;
            DlgStatic "Time Out:",WS_CHILD OR WS_VISIBLE,                                                   360,270,48 ,12 ,    IDC_PROMPT_MAIN_TIMEOUT
            DlgEdit WS_CHILD OR WS_VISIBLE OR WS_BORDER OR WS_TABSTOP OR ES_NUMBER OR 8000H,                424,268,48 ,12 ,    IDC_EDIT_MAIN_TIMEOUT
            DlgComCtl "msctls_updown32", UDS_AUTOBUDDY OR UDS_SETBUDDYINT OR UDS_ARROWKEYS \
                                                                   OR UDS_ALIGNRIGHT OR UDS_NOTHOUSANDS,    424,268,24 ,12 ,    IDC_SPINBOX_MAIN_TIMEOUT
            DlgCheck "Copy before 'Execute'",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,            360,282,130,12 ,    IDC_CHECK_MAIN_COPY_BEFORE
            DlgCheck "Ask before 'Copy before' Execute",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H, 360,296,130,12 ,    IDC_CHECK_MAIN_ASK_COPY_BEFORE
            DlgCheck "GUI on cancel ask 'Copy before'",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 3 OR 8000H,  360,310,130,12 ,    IDC_CHECK_MAIN_GUI_COPY_BEFORE
            ;
            ;===================================== BUTTON ===================================================================================== ;
            DlgButton "About...",WS_CHILD OR WS_TABSTOP OR 8000H,                                           012,338,88 ,16 ,    IDC_BUTTON_ABOUT
            ;DlgButton "Clarion Debugguer",WS_CHILD OR WS_TABSTOP OR WS_DISABLED OR 8000H,                   112,338,88 ,16 ,    IDC_BUTTON_CLADEBUG
            ;DlgButton "OutputDebugString Spy",WS_CHILD OR WS_TABSTOP  OR WS_DISABLED OR 8000H,                              212,338,88 ,16 ,    IDC_BUTTON_DEBUGSPY
            DlgButton "Save",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 8000H,                                 312,338,88 ,16 ,    IDC_BUTTON_SAVE
            DlgButton "Close",WS_CHILD OR WS_VISIBLE OR WS_TABSTOP OR 1 OR 8000H,                           412,338,88 ,16 ,    IDC_BUTTON_CLOSE
    ;
    CallModalDialog hInstance,0,MainDlgProc,ADDR lpArgs
    ;
    pop esi
    ;
    Invoke GlobalFree, lpArgs
    ;
    ret
    ;
main                            endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
;                       DIALOG PROCEDURE FOR WinMain
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
MainDlgProc                     proc hWnd:DWORD,uMsg:DWORD,wParam:DWORD,lParam:DWORD
                                ;
                                LOCAL hDummy    :DWORD
                                LOCAL hBtn      :HWND
                                LOCAL Rct       :RECT
                                LOCAL hBrush    :HBRUSH
                                LOCAL hOldBrush :HBRUSH
                                ;LOCAL icce      :INITCOMMONCONTROLSEX
        ;
        .if uMsg == WM_INITDIALOG
            ;
            mov eax,hWnd
            mov hWinMain,eax
            mov hOwnerPopMenu,eax
            ;
            Invoke SetWindowLong,hWnd,GWL_USERDATA,lParam
            mov eax, lParam
            mov eax, [eax]
            Invoke SendMessage,hWnd,WM_SETICON,1,[eax]
            mov eax, hWnd
            ;
            ;Put File Version in Windows Title
            Invoke GetFileVersionInfoSize,OFFSET szProgramPath,OFFSET dDummy
            .if (eax != 0)
                  push eax
                  Invoke GlobalAlloc,GPTR,eax
                  mov pMemory,eax
                  pop eax
                  Invoke GetFileVersionInfo,OFFSET szProgramPath,NULL,eax,pMemory
                  Invoke VerQueryValue,pMemory,OFFSET szFileVersion,OFFSET dDummy,OFFSET dwVersionLen
                  Invoke GetWindowText,hWnd,addr szBuffer,MAX_PATH
                  Invoke lstrcat,addr szBuffer,addr szVer
                  Invoke lstrcat,addr szBuffer,dDummy
                  Invoke lstrcat,addr szBuffer,addr szBracketOpen
                  Invoke lstrcat,addr szBuffer,addr szIniFile ;szProgramNameDotIni
                  Invoke lstrcat,addr szBuffer,addr szBracketClose
                  Invoke SetWindowText,hWnd,addr szBuffer
                  Invoke GlobalFree,pMemory
            .endif
            ;
            ;ICC_FLAGS = ICC_WIN95_CLASSES
            ;        ;comment out the ones you do not want
            ;        ICC_FLAGS = ICC_FLAGS or ICC_ANIMATE_CLASS
            ;        ICC_FLAGS = ICC_FLAGS or ICC_BAR_CLASSES
            ;        ICC_FLAGS = ICC_FLAGS or ICC_COOL_CLASSES
            ;        ICC_FLAGS = ICC_FLAGS or ICC_DATE_CLASSES
            ;        ICC_FLAGS = ICC_FLAGS or ICC_HOTKEY_CLASS
            ;        ICC_FLAGS = ICC_FLAGS or ICC_INTERNET_CLASSES
            ;        ICC_FLAGS = ICC_FLAGS or ICC_LISTVIEW_CLASSES
            ;        ICC_FLAGS = ICC_FLAGS or ICC_PAGESCROLLER_CLASS
            ;        ICC_FLAGS = ICC_FLAGS or ICC_PROGRESS_CLASS
            ;        ICC_FLAGS = ICC_FLAGS or ICC_TAB_CLASSES
            ;        ICC_FLAGS = ICC_FLAGS or ICC_TREEVIEW_CLASSES
            ;        ICC_FLAGS = ICC_FLAGS or ICC_UPDOWN_CLASS
            ;        ICC_FLAGS = ICC_FLAGS or ICC_USEREX_CLASSES
            ;mov icce.dwSize, SIZEOF INITCOMMONCONTROLSEX
            ;mov icce.dwICC,ICC_WIN95_CLASSES+ICC_LISTVIEW_CLASSES
            ;invoke InitCommonControlsEx,addr icce
            ;
            ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            ;*************** Apply Theme ******************
            ;
            xor eax,eax
            mov Rct.left,eax
            mov Rct.top,eax
            mov Rct.right,240
            mov Rct.bottom,60
            Invoke CreateFillBmp,FILL_NOISE,addr Rct,30,15,0,0,addr hBmpNoiseNormal,addr hBmpNoiseHover
            ;
            Invoke FillBuffer,addr BtnStruct, sizeof BtnStruct,0
            mov BtnStruct.hCursor_hover, NULL
            mov BtnStruct.hover_clr, $RGB (128, 0, 96)
            mov BtnStruct.push_clr, $RGB (128, 0, 96)
            mov BtnStruct.normal_clr, $RGB (0, 0, 64)
            ;STYLE_DEFAULT,STYLE_XNET,STYLE_OFFICE_XP,STYLE_OFFICE_2000
            ;mov BtnStruct.btn_prop, STYLE_DEFAULT+FLAT_BTN+TRANSPARENT_BKGND
            mov BtnStruct.btn_prop, STYLE_XNET+FLAT_BTN+DRAW_BORDER+DRAW_FOCUS+TRANSPARENT_BKGND
            ;mov BtnStruct.btn_prop, STYLE_OFFICE_XP+FLAT_BTN+TRANSPARENT_BKGND
            ;mov BtnStruct.btn_prop, STYLE_OFFICE_2000+FLAT_BTN+TRANSPARENT_BKGND
            ;mov BtnStruct.btn_prop, FLAT_BTN+TRANSPARENT_BKGND+RGN_BUTTON ;Sєlo texto
            ;mov BtnStruct.btn_prop, FLAT_BTN+STYLE_XNET+TRANSPARENT_BKGND
            ;mov BtnStruct.btn_prop, FLAT_BTN+STYLE_XNET
            ;mov BtnStruct.btn_prop, FLAT_BTN+DRAW_BORDER+STYLE_XNET+DRAW_FOCUS+TRANSPARENT_BKGND+RGN_BUTTON
            ;mov BtnStruct.btn_prop, FLAT_BTN+DRAW_BORDER+DRAW_FOCUS+TRANSPARENT_BKGND+RGN_BUTTON ;Igual que el fondo
            ;mov BtnStruct.btn_prop, FLAT_BTN+DRAW_BORDER+DRAW_FOCUS+TRANSPARENT_BKGND
            ;mov BtnStruct.btn_prop, FLAT_BTN+DRAW_FOCUS+TRANSPARENT_BKGND+RGN_BUTTON
            ;xor eax,eax
            ;mov BtnStruct.hIcon_normal,eax
            ;mov BtnStruct.hIcon_hover,eax
            mov ecx,IDC_BUTTON_ABOUT
            @@:
                push ecx
                mov hDummy,$invoke(GetDlgItem,hWnd,ecx)
                Invoke RedrawButton,hDummy,addr BtnStruct
                pop ecx
                add ecx,1
                cmp ecx,IDC_BUTTON_CLOSE
                jle @B

        ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            ;
            ;*************** Fills up TimerMinutes ******************
            ;
            ;
            Invoke  GetDlgItem, hWnd, IDC_SPINBOX_PRE_TIMEOUT
            Invoke  SendMessage, eax, UDM_SETRANGE, 0, 00007FFFH   ;wordLower,wordUpper 0,99
            ;
            ;*************** Fills up TimerSeconds ******************
            ;
            ;
            Invoke  GetDlgItem, hWnd, IDC_SPINBOX_MAIN_TIMEOUT
            Invoke  SendMessage, eax, UDM_SETRANGE, 0, 00007FFFH   ;wordLower,wordUpper 0,59
            ;
            ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                       ;Read Values from Ini File
            ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            ;
            Invoke ReadIniFile,hWnd
            ;
            ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                       ;Enable Checks
            ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            ;
            Invoke IsDlgButtonChecked,hWnd,IDC_CHECK_PRE_UNHIDE
            .if eax==BST_CHECKED
                Invoke GetDlgItem,hWnd,IDC_CHECK_PRE_TERM
                Invoke EnableWindow,eax,FALSE
            .elseif eax==BST_UNCHECKED
                Invoke GetDlgItem,hWnd,IDC_CHECK_PRE_TERM
                Invoke EnableWindow,eax,TRUE
            .endif
            ;
            ;
            Invoke IsDlgButtonChecked,hWnd,IDC_CHECK_MAIN_UNHIDE
            .if eax==BST_CHECKED
                Invoke GetDlgItem,hWnd,IDC_CHECK_MAIN_TERM
                Invoke EnableWindow,eax,FALSE
            .elseif eax==BST_UNCHECKED
                Invoke GetDlgItem,hWnd,IDC_CHECK_MAIN_TERM
                Invoke EnableWindow,eax,TRUE
            .endif
            ;
            ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                                       ;CREATE TOOLTIP
            ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            Invoke CreateWindowEx,NULL,addr szToolTipClass,NULL,\
                                    WS_POPUP OR TTS_NOPREFIX OR TTS_ALWAYSTIP OR TTS_BALLOON\
                                    ,0,0,0,0,hWnd,NULL,hInstance,NULL
            mov    hToolTip,eax
            Invoke SetWindowPos,hToolTip,HWND_TOPMOST,0,0,0,0,SWP_NOMOVE or SWP_NOSIZE or SWP_NOACTIVATE
            Invoke SendMessage,hToolTip,TTM_ACTIVATE,TRUE,0
            Invoke SendMessage,hToolTip,TTM_SETMAXTIPWIDTH,0, 300 ;It enables a multiline tooltip by setting the display rectangle to 150 pixels.
            ;The text buffer specified by the szText member of the NMTTDISPINFO structure can accommodate only 80 characters.
            ;   If you need to use a longer string, point the lpszText member of NMTTDISPINFO to a buffer containing the desired text.
            ; I think the tooltip-ctrl is able to handle multiline tooltips only from commctrl.dll version 4.71
            ;
            ;INITIALIZE MEMBERS OF THE TOOLINFO STRUCTURE
            mov    ti.cbSize,sizeof TOOLINFO
            mov    ti.uFlags,TTF_SUBCLASS or TTF_IDISHWND
            push   hWnd
            pop    ti.hWnd
            push   hInstance
            pop    ti.hInst
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDIT_PRE_SOURCE
            Invoke GetDlgItem, hWnd,IDC_EDIT_PRE_SOURCE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt00
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_MAIN_SOURCE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt00
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_BUTTON_PRE_SOURCE_OVERWR
            Invoke GetDlgItem, hWnd,IDC_BUTTON_PRE_SOURCE_OVERWR
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt1
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_BUTTON_MAIN_SOURCE_OVERWR
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt1
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDIT_PRE_TARGET
            Invoke GetDlgItem, hWnd,IDC_EDIT_PRE_TARGET
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt0
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_MAIN_TARGET
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt0
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_BUTTON_PRE_TARGET_OVERWR
            Invoke GetDlgItem, hWnd,IDC_BUTTON_PRE_TARGET_OVERWR
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt2
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_BUTTON_MAIN_TARGET_OVERWR
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt2
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDIT_PRE_EXECUTE
            Invoke GetDlgItem, hWnd,IDC_EDIT_PRE_EXECUTE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt0
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_MAIN_EXECUTE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt0
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDIT_PRE_PARAMS
            Invoke GetDlgItem, hWnd,IDC_EDIT_PRE_PARAMS
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt0
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_MAIN_PARAMS
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt0
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDIT_PRE_WRKDIR
            Invoke GetDlgItem, hWnd,IDC_EDIT_PRE_WRKDIR
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt0
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_MAIN_WRKDIR
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt0
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHECK_PRE_UNHIDE
            Invoke GetDlgItem, hWnd,IDC_CHECK_PRE_UNHIDE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt5
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_CHECK_MAIN_UNHIDE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt5
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHECK_PRE_SHOW
            Invoke GetDlgItem, hWnd,IDC_CHECK_PRE_SHOW
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt3
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_CHECK_MAIN_SHOW
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt3
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHECK_PRE_ASK_BEFORE_EXEC
            Invoke GetDlgItem, hWnd,IDC_CHECK_PRE_ASK_BEFORE_EXEC
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt8
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_CHECK_MAIN_ASK_BEFORE_EXEC
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt8
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHECK_PRE_GUI_CANCEL_EXEC
            Invoke GetDlgItem, hWnd,IDC_CHECK_PRE_GUI_CANCEL_EXEC
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxtA
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_CHECK_MAIN_GUI_CANCEL_EXEC
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxtA
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHECK_PRE_TERM
            Invoke GetDlgItem, hWnd,IDC_CHECK_PRE_TERM
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt6
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_CHECK_MAIN_TERM
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt6
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHECK_PRE_COPY_AFTER
            Invoke GetDlgItem, hWnd,IDC_CHECK_PRE_COPY_AFTER
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxtB
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_CHECK_MAIN_COPY_AFTER
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxtB
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHECK_PRE_ASK_COPY_AFTER
            Invoke GetDlgItem, hWnd,IDC_CHECK_PRE_ASK_COPY_AFTER
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt7
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_CHECK_MAIN_ASK_COPY_AFTER
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt7
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHECK_PRE_GUI_COPY_AFTER
            Invoke GetDlgItem, hWnd,IDC_CHECK_PRE_GUI_COPY_AFTER
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt9
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_CHECK_MAIN_GUI_COPY_AFTER
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt9
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_EDIT_PRE_TIMEOUT
            Invoke GetDlgItem, hWnd,IDC_EDIT_PRE_TIMEOUT
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt4
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_MAIN_TIMEOUT
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt4
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHECK_PRE_COPY_BEFORE
            Invoke GetDlgItem, hWnd,IDC_CHECK_PRE_COPY_BEFORE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxtC
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_CHECK_MAIN_COPY_BEFORE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxtC
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHECK_PRE_ASK_COPY_BEFORE
            Invoke GetDlgItem, hWnd,IDC_CHECK_PRE_ASK_COPY_BEFORE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt7
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_CHECK_MAIN_ASK_COPY_BEFORE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt7
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_CHECK_PRE_GUI_COPY_BEFORE
            Invoke GetDlgItem, hWnd,IDC_CHECK_PRE_GUI_COPY_BEFORE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt9
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_CHECK_MAIN_GUI_COPY_BEFORE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxt9
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            ;
            ;SEND AN ADDTOOL MESSAGE TO THE TOOLTIP CONTROL WINDOW: IDC_BUTTON_SAVE
            Invoke GetDlgItem, hWnd,IDC_BUTTON_SAVE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxtD
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;
            Invoke GetDlgItem, hWnd,IDC_BUTTON_CLOSE
            mov    ti.uId,eax
            mov    ti.lpszText,offset szToolTxtE
            Invoke SendMessage,hToolTip,TTM_ADDTOOL,0,addr ti
            ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            ;
            ;*************** Create PopUpMenu *******************
            ;Invoke RegisterHotKey, hWnd, EQU_HOTKEY,MOD_CONTROL, VK_SPACE
            Invoke CreatePopupMenu        ; Create a PopupMenu structure
            mov hPopupMenu, eax               ; and save it's pointer
            ;lazy code below to add items to the menu used when right clicking

            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_COMPUTERNAME, addr szCOMPUTERNAME
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_DATE, addr szDATE
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_DAY, addr szDAY
            ;Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_DRIVE, addr szDRIVE
            ;Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_ERRORCODE, addr szERRORCODE
            ;Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_ERRORTEXT, addr szERRORTEXT
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_HOUR, addr szHOUR
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_MINUTE, addr szMINUTE
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_MONTH, addr szMONTH
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_MYDOCUMENTS, addr szMYDOCUMENTS
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_PROCESSID, addr szPROCESSID
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_PROGRAMDIR, addr szPROGRAMDIR
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_PROGRAMFILES, addr szPROGRAMFILES
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_PROGRAMNAME, addr szPROGRAMNAME
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_SECOND, addr szSECOND
            ;Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_SHARE, addr szSHARE
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_SYSTEMDIR, addr szSYSTEMDIR
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_SYSTEMDRIVE, addr szSYSTEMDRIVE
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_TEMPDIR, addr szTEMPDIR
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_TIME, addr szTIME
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_USERNAME, addr szUSERNAME
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_WEEKDAY, addr szWEEKDAY
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_WINDIR, addr szWINDIR
            Invoke AppendMenu, hPopupMenu, MF_STRING,IDM_YEAR, addr szYEAR
            ;
            ;Invoke AppendMenu, hPopupMenu, MF_SEPARATOR, 0, 0 ; <-- dress it up ;)
            ;Invoke AppendMenu, mnuhWnd, MF_STRING or MF_GRAYED, IDM_SIX, addr Menu6
            ;Invoke AppendMenu, mnuhWnd, MF_STRING or MF_CHECKED, IDM_SEVEN,addr Menu7
            ;Invoke  LoadBitmap, hInstance, MYBITMAP  ; load a .bmp to display
            ;mov     hBMP,eax                         ; and store it's handle
            ;Invoke AppendMenu, mnuhWnd, MF_BITMAP, IDM_EIGHT, hBMP ; a bmp !!
            ;and you can keep going up to your screen size limit
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_PRE_SOURCE
            mov ActiveIDC[0].hIDC,eax
            Invoke SetWindowLong,eax,GWL_WNDPROC,addr SubWndProc
            mov ActiveIDC[0].hProc,eax
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_PRE_TARGET
            mov ActiveIDC[8].hIDC,eax
            Invoke SetWindowLong,eax,GWL_WNDPROC,addr SubWndProc
            mov ActiveIDC[8].hProc,eax
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_PRE_EXECUTE
            mov ActiveIDC[16].hIDC,eax
            Invoke SetWindowLong,eax,GWL_WNDPROC,addr SubWndProc
            mov ActiveIDC[16].hProc,eax
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_PRE_PARAMS
            mov ActiveIDC[24].hIDC,eax
            Invoke SetWindowLong,eax,GWL_WNDPROC,addr SubWndProc
            mov ActiveIDC[24].hProc,eax
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_PRE_WRKDIR
            mov ActiveIDC[32].hIDC,eax
            Invoke SetWindowLong,eax,GWL_WNDPROC,addr SubWndProc
            mov ActiveIDC[32].hProc,eax
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_MAIN_SOURCE
            mov ActiveIDC[40].hIDC,eax
            Invoke SetWindowLong,eax,GWL_WNDPROC,addr SubWndProc
            mov ActiveIDC[40].hProc,eax
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_MAIN_TARGET
            mov ActiveIDC[48].hIDC,eax
            Invoke SetWindowLong,eax,GWL_WNDPROC,addr SubWndProc
            mov ActiveIDC[48].hProc,eax
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_MAIN_EXECUTE
            mov ActiveIDC[56].hIDC,eax
            Invoke SetWindowLong,eax,GWL_WNDPROC,addr SubWndProc
            mov ActiveIDC[56].hProc,eax
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_MAIN_PARAMS
            mov ActiveIDC[64].hIDC,eax
            Invoke SetWindowLong,eax,GWL_WNDPROC,addr SubWndProc
            mov ActiveIDC[64].hProc,eax
            ;
            Invoke GetDlgItem, hWnd,IDC_EDIT_MAIN_WRKDIR
            mov ActiveIDC[72].hIDC,eax
            Invoke SetWindowLong,eax,GWL_WNDPROC,addr SubWndProc
            mov ActiveIDC[72].hProc,eax
            ;
            ;********* PUT This Windows ABOVE it all
            Invoke ForceWindowToTop,hWnd
            ;
        ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        .elseif uMsg == WM_ERASEBKGND
            invoke GetClientRect, hWnd, addr Rct
            mov hBrush, $invoke (CreatePatternBrush,hBmpNoiseNormal)
            invoke SelectObject, wParam, hBrush
            mov hOldBrush, eax
            invoke PatBlt, wParam, Rct.left, Rct.top, Rct.right, Rct.bottom, PATCOPY
            invoke SelectObject, wParam, hOldBrush
            invoke DeleteObject, hBrush
            return TRUE
        .elseif uMsg==WM_WINDOWPOSCHANGED
            Invoke GetWindowRect, hWnd,addr WinMainRec
            Invoke RedrawWindow,0,addr WinMainRec,0,RDW_INVALIDATE OR RDW_ALLCHILDREN OR RDW_UPDATENOW
        .elseif uMsg == WM_CTLCOLORSTATIC
              Invoke GetDlgCtrlID, lParam
              Invoke SetBkMode,wParam,TRANSPARENT ;OPAQUE ;TRANSPARENT ;TRANSPARENT=Mancha(No actualiza)
              ;INVOKE SetBkColor,wParam,hBmpNoiseNormal
              Invoke CreatePatternBrush,hBmpNoiseNormal
              ret
        .elseif uMsg == WM_COMMAND
            ;
            mov eax,wParam
            ;
            .if (wParam == IDC_CHECK_PRE_UNHIDE)
                ;********************* Check Unhide Previous Task*************
                shr eax,16
                .if ax==BN_CLICKED
                    ;
                    Invoke IsDlgButtonChecked,hWnd,IDC_CHECK_PRE_UNHIDE
                    .if eax==BST_CHECKED
                        Invoke GetDlgItem,hWnd,IDC_CHECK_PRE_TERM
                        Invoke EnableWindow,eax,FALSE
                    .elseif eax==BST_UNCHECKED
                        Invoke GetDlgItem,hWnd,IDC_CHECK_PRE_TERM
                        Invoke EnableWindow,eax,TRUE
                    .endif
                    ;
                .endif
            .elseif (wParam == IDC_CHECK_MAIN_UNHIDE)
                ;********************* Check Unhide Main Task*************
                shr eax,16
                .if ax==BN_CLICKED
                    ;
                    Invoke IsDlgButtonChecked,hWnd,IDC_CHECK_MAIN_UNHIDE
                    .if eax==BST_CHECKED
                        Invoke GetDlgItem,hWnd,IDC_CHECK_MAIN_TERM
                        Invoke EnableWindow,eax,FALSE
                    .elseif eax==BST_UNCHECKED
                        Invoke GetDlgItem,hWnd,IDC_CHECK_MAIN_TERM
                        Invoke EnableWindow,eax,TRUE
                    .endif
                    ;
                .endif
            .elseif     (wParam == IDC_BUTTON_PRE_SOURCE_FILEDLG) || (wParam == IDC_BUTTON_MAIN_SOURCE_FILEDLG)
                ;********************* Copy Source *************
                shr eax,16
                .if ax==BN_CLICKED
                    .if (wParam == IDC_BUTTON_PRE_SOURCE_FILEDLG)
                        Invoke FillBuffer,addr szPreCopySource,sizeof szPreCopySource,0
                        Invoke MessageBox,hWnd,addr szAskForFileText,addr szAskForFileTitle,MB_YESNO+MB_ICONQUESTION
                        .if eax != IDNO
                            Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szPreCopySource,sizeof szPreCopySource,addr szFileDialogFolder,sizeof szFileDialogFolder,TRUE
                        .else
                            Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szPreCopySource,sizeof szPreCopySource,addr szFileDialogNoFilter,sizeof szFileDialogNoFilter,FALSE
                        .endif
                        .if eax!=0
                            Invoke SetDlgItemText,hWnd,IDC_EDIT_PRE_SOURCE,addr szPreCopySource
                        .endif
                    .else
                        Invoke FillBuffer,addr szCopySource,sizeof szCopySource,0
                        Invoke MessageBox,hWnd,addr szAskForFileText,addr szAskForFileTitle,MB_YESNO+MB_ICONQUESTION
                        .if eax != IDNO
                            Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szCopySource,sizeof szCopySource,addr szFileDialogFolder,sizeof szFileDialogFolder,TRUE
                        .else
                            Invoke FileDialog,hWnd,offset szFileDialogTitle,offset szCopySource,sizeof szCopySource,offset szFileDialogNoFilter,sizeof szFileDialogNoFilter,FALSE
                        .endif
                        .if eax!=0
                            Invoke SetDlgItemText,hWnd,IDC_EDIT_MAIN_SOURCE,addr szCopySource
                        .endif
                    .endif
                .endif
            .elseif (wParam == IDC_BUTTON_PRE_TARGET_FILEDLG) || (wParam == IDC_BUTTON_MAIN_TARGET_FILEDLG)
                ;********************* Copy Target *************
                shr eax,16
                .if ax==BN_CLICKED
                    .if (wParam == IDC_BUTTON_PRE_TARGET_FILEDLG)
                        Invoke FillBuffer,addr szPreCopyTarget,sizeof szPreCopyTarget,0
                        Invoke MessageBox,hWnd,addr szAskForFileText,addr szAskForFileTitle,MB_YESNO+MB_ICONQUESTION
                        .if eax != IDNO
                            Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szPreCopyTarget,sizeof szPreCopyTarget,addr szFileDialogFolder,sizeof szFileDialogFolder,TRUE
                        .else
                            Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szPreCopyTarget,sizeof szPreCopyTarget,addr szFileDialogNoFilter,sizeof szFileDialogNoFilter,FALSE
                        .endif
                        .if eax!=0
                            Invoke SetDlgItemText,hWnd,IDC_EDIT_PRE_TARGET,addr szPreCopyTarget
                        .endif
                    .else
                        Invoke FillBuffer,addr szCopyTarget,sizeof szCopyTarget,0
                        Invoke MessageBox,hWnd,addr szAskForFileText,addr szAskForFileTitle,MB_YESNO+MB_ICONQUESTION
                        .if eax != IDNO
                            Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szCopyTarget,sizeof szCopyTarget,addr szFileDialogFolder,sizeof szFileDialogFolder,TRUE
                        .else
                            Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szCopyTarget,sizeof szCopyTarget,addr szFileDialogNoFilter,sizeof szFileDialogNoFilter,FALSE
                        .endif
                        .if eax!=0
                            Invoke SetDlgItemText,hWnd,IDC_EDIT_MAIN_TARGET,addr szCopyTarget
                        .endif
                    .endif
                .endif
            .elseif (wParam == IDC_BUTTON_PRE_EXECUTE_FILEDLG) || (wParam == IDC_BUTTON_MAIN_EXECUTE_FILEDLG)
                ;********************* Execute *************
                shr eax,16
                .if ax==BN_CLICKED
                    .if (wParam == IDC_BUTTON_PRE_EXECUTE_FILEDLG)
                        Invoke FillBuffer,addr szPreCmdRunProgram,sizeof szPreCmdRunProgram,0
                        Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szPreCmdRunProgram,sizeof szPreCmdRunProgram,addr szFileDialogNoFilter,sizeof szFileDialogNoFilter,FALSE
                        .if eax!=0
                            Invoke SetDlgItemText,hWnd,IDC_EDIT_PRE_EXECUTE,addr szPreCmdRunProgram
                        .endif
                    .else
                        Invoke FillBuffer,addr szCmdRunProgram,sizeof szCmdRunProgram,0
                        Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szCmdRunProgram,sizeof szCmdRunProgram,addr szFileDialogNoFilter,sizeof szFileDialogNoFilter,FALSE
                        .if eax!=0
                            Invoke SetDlgItemText,hWnd,IDC_EDIT_MAIN_EXECUTE,addr szCmdRunProgram
                        .endif
                    .endif
                .endif
            .elseif (wParam == IDC_BUTTON_PRE_PARAMS_FILEDLG) || (wParam == IDC_BUTTON_MAIN_PARAMS_FILEDLG)
                ;********************* Parameters *************
                shr eax,16
                .if ax==BN_CLICKED
                    .if (wParam == IDC_BUTTON_PRE_PARAMS_FILEDLG)
                        Invoke FillBuffer,addr szPreCmdRunParameters,sizeof szPreCmdRunParameters,0
                        Invoke MessageBox,hWnd,addr szAskForFileText,addr szAskForFileTitle,MB_YESNO+MB_ICONQUESTION
                        .if eax != IDNO
                            Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szPreCmdRunParameters,sizeof szPreCmdRunParameters,addr szFileDialogFolder,sizeof szFileDialogFolder,TRUE
                        .else
                            Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szPreCmdRunParameters,sizeof szPreCmdRunParameters,addr szFileDialogNoFilter,sizeof szFileDialogNoFilter,FALSE
                        .endif
                        .if eax!=0
                            Invoke SetDlgItemText,hWnd,IDC_EDIT_PRE_PARAMS,addr szPreCmdRunParameters
                        .endif
                    .else
                        Invoke FillBuffer,addr szCmdRunParameters,sizeof szCmdRunParameters,0
                        Invoke MessageBox,hWnd,addr szAskForFileText,addr szAskForFileTitle,MB_YESNO+MB_ICONQUESTION
                        .if eax != IDNO
                            Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szCmdRunParameters,sizeof szCmdRunParameters,addr szFileDialogFolder,sizeof szFileDialogFolder,TRUE
                        .else
                            Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szCmdRunParameters,sizeof szCmdRunParameters,addr szFileDialogNoFilter,sizeof szFileDialogNoFilter,FALSE
                        .endif
                        .if eax!=0
                            Invoke SetDlgItemText,hWnd,IDC_EDIT_MAIN_PARAMS,addr szCmdRunParameters
                        .endif
                    .endif
                .endif
                ;
            .elseif (wParam == IDC_BUTTON_PRE_WRKDIR_FILEDLG) || (wParam == IDC_BUTTON_MAIN_WRKDIR_FILEDLG)
                ;********************* Parameters *************
                shr eax,16
                .if ax==BN_CLICKED
                    .if (wParam == IDC_BUTTON_PRE_WRKDIR_FILEDLG)
                        Invoke FillBuffer,addr szPreCmdRunWorkingDir,sizeof szPreCmdRunWorkingDir,0
                        Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szPreCmdRunWorkingDir,sizeof szPreCmdRunWorkingDir,addr szFileDialogFolder,sizeof szFileDialogFolder,TRUE
                        .if eax!=0
                            Invoke SetDlgItemText,hWnd,IDC_EDIT_PRE_WRKDIR,addr szPreCmdRunWorkingDir
                        .endif
                    .else
                        Invoke FillBuffer,addr szCmdRunWorkingDir,sizeof szCmdRunWorkingDir,0
                        Invoke FileDialog,hWnd,addr szFileDialogTitle,addr szCmdRunWorkingDir,sizeof szCmdRunWorkingDir,addr szFileDialogFolder,sizeof szFileDialogFolder,TRUE
                        .if eax!=0
                            Invoke SetDlgItemText,hWnd,IDC_EDIT_MAIN_WRKDIR,addr szCmdRunWorkingDir
                        .endif
                    .endif
                .endif
                ;
            .elseif wParam == IDC_BUTTON_ABOUT     ;About
                shr eax,16
                .if ax==BN_CLICKED
                    .if bIsClarionDebugRunning
                        Invoke AboutDlg,hWnd,hIconClaDebug
                    .else
                        Invoke AboutDlg,hWnd,hIcon
                    .endif
                .endif
            .elseif wParam == IDC_BUTTON_SAVE     ;Save
                shr eax,16
                .if ax==BN_CLICKED
                    Invoke WriteIniFile,hWnd
                    .if eax==0 ;Returns 0 if fails. No permissions to write on folder or ReadOnly ini file?
                        Invoke GetLastError
                        mov ecx,eax
                        Invoke FormatMessage,FORMAT_MESSAGE_FROM_SYSTEM, 0,ecx, 0,addr szCmdCompare,sizeof szCmdCompare, 0
                        .if eax!=0
                            lea ecx,szCmdCompare
                            add ecx,eax
                            mov byte ptr[ecx],10
                            inc eax
                            mov byte ptr[ecx],0
                            Invoke  lstrcat,addr szCmdCompare,addr szIniFile
                            Invoke MessageBox,NULL,addr szCmdCompare,addr szErrorSavingIniFileTitle,MB_OK+MB_ICONHAND+MB_SYSTEMMODAL
                        .endif
                    .endif
                .endif
            .elseif wParam == IDC_BUTTON_CLOSE     ;Close
                shr eax,16
                .if ax==BN_CLICKED
                    jmp ExitMainDlgProc
                .endif
            .endif
        .elseif uMsg == WM_CLOSE
            ExitMainDlgProc:
            Invoke EndDialog,hWnd,0
        .endif
        xor eax, eax
    ret
MainDlgProc                     endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
;                       About Dialog
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
AboutDlg                        proc hParent:DWORD,lpIcon:DWORD
                                ; OR SS_WHITERECT OR SS_SUNKEN OR SS_WHITEFRAME OR SS_WHITEFRAME OR SS_GRAYRECT
                                ; OR SS_ETCHEDVERT OR SS_ETCHEDHORZ OR SS_ETCHEDFRAME
    Dialog "About RunHide","MS Sans Serif",8, \            ; caption,font,pointsize
            WS_OVERLAPPED or WS_SYSMENU or DS_CENTER, \     ; style
            3, \                                            ; control count
            0,0,304,324, \                                  ; x y co-ordinates
            1024                                            ; memory buffer size
    DlgButton " Boogie ",BS_GROUPBOX  OR BS_FLAT OR CCS_NODIVIDER OR BS_CENTER OR SS_ETCHEDVERT,8,5,283,266,IDGROUP3
    DlgStatic "[ www.carabez.com ]",WS_CHILD OR WS_VISIBLE OR SS_CENTER OR SS_NOTIFY,117,266,70,12,IDSTATIC12
    DlgButton "OK",WS_TABSTOP OR 1 OR 8000H OR SS_ETCHEDVERT,117,282,70,16,IDC_BTN6


    CallModalDialog hInstance,hParent,AboutDlgProc,ADDR hParent

    ret

AboutDlg                        endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
;                       AboutDlg Procedure
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
AboutDlgProc                    proc hWin:DWORD,uMsg:DWORD,wParam:DWORD,lParam:DWORD
                                LOCAL hDummy    :DWORD
                                LOCAL hBtn      :HWND
                                LOCAL Rct       :RECT
                                LOCAL hBrush    :HBRUSH
                                LOCAL hOldBrush :HBRUSH
    ;
    .if uMsg == WM_INITDIALOG
        ; -----------------------------------------
        ; write the parameters passed in "lParam"
        ; to the dialog's GWL_USERDATA address.
        ; -----------------------------------------
        Invoke SetWindowLong,hWin,GWL_USERDATA,lParam
        mov eax, lParam
        mov eax, [eax+4]
        Invoke SendMessage,hWin,WM_SETICON,1,eax
        ;

        ;hOldStaticProc
        ;Load bitmap from resource
        Invoke BitmapFromResource, hInstance, 600
        mov hBmp, eax
        ;
        Invoke  ConvertStaticToHyperlink,hWin,IDSTATIC12
        ;Put File Version in Windows Title
        Invoke GetFileVersionInfoSize,OFFSET szProgramPath,OFFSET dDummy
        .if (eax != 0)
              push eax
              Invoke GlobalAlloc,GPTR,eax
              mov pMemory,eax
              pop eax
              Invoke GetFileVersionInfo,OFFSET szProgramPath,NULL,eax,pMemory
              Invoke VerQueryValue,pMemory,OFFSET szFileVersion,OFFSET dDummy,OFFSET dwVersionLen
              Invoke GetWindowText,hWin,addr szBuffer,MAX_PATH
              Invoke lstrcat,addr szBuffer,addr szVer
              Invoke lstrcat,addr szBuffer,dDummy
              Invoke SetWindowText,hWin,addr szBuffer
              Invoke GlobalFree,pMemory
        .endif
    ;>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    .elseif uMsg == WM_ERASEBKGND
        invoke GetClientRect, hWin, addr Rct
        mov hBrush, $invoke (CreatePatternBrush,hBmpNoiseNormal)
        invoke SelectObject, wParam, hBrush
        mov hOldBrush, eax
        invoke PatBlt, wParam, Rct.left, Rct.top, Rct.right, Rct.bottom, PATCOPY
        invoke SelectObject, wParam, hOldBrush
        invoke DeleteObject, hBrush
        return TRUE
    .elseif uMsg==WM_WINDOWPOSCHANGED
        Invoke GetWindowRect,hWin,addr WinMainRec
        Invoke RedrawWindow,0,addr WinMainRec,0,RDW_INVALIDATE OR RDW_ALLCHILDREN OR RDW_UPDATENOW
    .elseif uMsg == WM_PAINT
        ;Invoke Paint_Proc,hWin,50,40,300,225
        Invoke Paint_Proc,hWin,24,26,400,400
    .elseif uMsg == WM_CTLCOLORSTATIC
        ;wParam = Handle to the device context for the static control window.
        ;lParam = Handle to the static control.
        Invoke GetDlgCtrlID, lParam
        ;.if lParam!=IDSTATIC12
        Invoke SetBkMode,wParam,TRANSPARENT ;OPAQUE ;TRANSPARENT ;TRANSPARENT=Mancha(No actualiza)
        ;.endif
        ;INVOKE SetBkColor,wParam,hBmpNoiseNormal
        Invoke CreatePatternBrush,hBmpNoiseNormal
        ret
    .elseif uMsg == WM_COMMAND
          .if wParam == IDC_BTN6
              jmp Exit_About
          .elseif wParam == IDSTATIC12
              Invoke ShellExecute,hWin,NULL,addr szWebPage,NULL,NULL,SW_SHOWDEFAULT
          .endif
    .elseif uMsg == WM_CLOSE
        Exit_About:
        Invoke EndDialog,hWin,0
    .endif
    ;
    xor eax, eax
    ret

AboutDlgProc                    endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;Paint Procedure
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
Paint_Proc                      proc hWin:DWORD, xPos:DWORD, yPos:DWORD, wPos:DWORD, hPos:DWORD
                        ;
                        LOCAL hDC   :DWORD
                        LOCAL hOld  :DWORD
                        LOCAL memDC :DWORD
                        LOCAL ps    :PAINTSTRUCT
      ;
      Invoke BeginPaint,hWin,ADDR ps
      mov hDC, eax
      Invoke CreateCompatibleDC,hDC
      mov memDC, eax
      Invoke SelectObject,memDC,hBmp
      mov hOld, eax
      Invoke BitBlt,hDC,xPos,yPos,wPos,hPos,memDC,0,0,SRCCOPY
      Invoke SelectObject,hDC,hOld
      Invoke DeleteDC,memDC
      Invoke EndPaint,hWin,ADDR ps
      ;
      ret
      ;
Paint_Proc                      endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;HyperLink Relative
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
OPTION PROLOGUE:NONE
OPTION EPILOGUE:NONE
align 4
_HyperlinkParentProc            proc hWin:DWORD,uMsg:DWORD,wParam:DWORD,lParam:DWORD
      thisPROC_hWnd TEXTEQU <[esp+4*1]>
      thisPROC_uMsg TEXTEQU <[esp+4*2]>
      thisPROC_wParam TEXTEQU <[esp+4*3]>
      thisPROC_lParam TEXTEQU <[esp+4*4]>

      invoke  GetProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_PROC
      mov ecx, thisPROC_uMsg
      pop edx
      push  eax
      cmp ecx, WM_CTLCOLORSTATIC
      push  edx ; return address
      je  __WM_CTLCOLORSTATIC
      cmp ecx, WM_DESTROY
      je  __WM_DESTROY
      jmp CallWindowProc

      ; another dword on the stack
      thisPROC_old  TEXTEQU <[4][esp+4*0]>
      thisPROC_hWnd TEXTEQU <[4][esp+4*1]>
      thisPROC_uMsg TEXTEQU <[4][esp+4*2]>
      thisPROC_wParam TEXTEQU <[4][esp+4*3]>
      thisPROC_lParam TEXTEQU <[4][esp+4*4]>

      __WM_DESTROY:
      invoke  SetWindowLong, thisPROC_hWnd[8], GWL_WNDPROC, eax
      invoke  RemoveProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_PROC
      _x: jmp CallWindowProc


      __WM_CTLCOLORSTATIC:
      ; why would another control have this property?
      invoke  GetProp, thisPROC_lParam[4],addr PROP_ORIGINAL_PROC
      test  eax, eax
      je  _x
      invoke  CallWindowProc, thisPROC_old[16], thisPROC_hWnd[12], thisPROC_uMsg[8], thisPROC_wParam[4], thisPROC_lParam
      mov [esp+4], eax
      invoke  SetTextColor, thisPROC_wParam[4],00C40000H ;CONST_RGB ;(0,0,192)
      mov eax, [esp+4]
      retn 20
_HyperlinkParentProc            endp
OPTION PROLOGUE:PROLOGUEDEF
OPTION EPILOGUE:EPILOGUEDEF
;
OPTION PROLOGUE:NONE
OPTION EPILOGUE:NONE
align 4
_HyperlinkProc                  proc hWin:DWORD,uMsg:DWORD,wParam:DWORD,lParam:DWORD
      thisPROC_hWnd TEXTEQU <[esp+4*1]>
      thisPROC_uMsg TEXTEQU <[esp+4*2]>
      thisPROC_wParam TEXTEQU <[esp+4*3]>
      thisPROC_lParam TEXTEQU <[esp+4*4]>

      invoke  GetProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_PROC
      mov ecx, thisPROC_uMsg
      pop edx
      push  eax
      cmp ecx, WM_MOUSEMOVE
      push  edx ; return address
      je  __WM_MOUSEMOVE
      cmp ecx, WM_SETCURSOR
      je  __WM_SETCURSOR
      cmp ecx, WM_DESTROY
      je  __WM_DESTROY
      jmp CallWindowProc

      ; another dword on the stack
      thisPROC_old  TEXTEQU <[4][esp+4*0]>
      thisPROC_hWnd TEXTEQU <[4][esp+4*1]>
      thisPROC_uMsg TEXTEQU <[4][esp+4*2]>
      thisPROC_wParam TEXTEQU <[4][esp+4*3]>
      thisPROC_lParam TEXTEQU <[4][esp+4*4]>

      __WM_DESTROY:
      invoke  SetWindowLong, thisPROC_hWnd[8], GWL_WNDPROC, eax
      invoke  RemoveProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_PROC

      invoke  GetProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_FONT
      invoke  SendMessage, thisPROC_hWnd[12], WM_SETFONT, eax, 0
      invoke  RemoveProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_FONT

      invoke  GetProp, thisPROC_hWnd[4],addr PROP_UNDERLINE_FONT
      invoke  DeleteObject, eax
      invoke  RemoveProp, thisPROC_hWnd[4],addr PROP_UNDERLINE_FONT
      jmp CallWindowProc


      __WM_MOUSEMOVE:
      invoke  GetCapture
      cmp eax, thisPROC_hWnd
      jne _start_capture

      sub esp, SIZEOF RECT
      invoke  GetWindowRect, thisPROC_hWnd[4 + SIZEOF RECT], esp
      movsx eax, WORD PTR thisPROC_lParam[2 + SIZEOF RECT]
      movsx edx, WORD PTR thisPROC_lParam[0 + SIZEOF RECT]
      push  eax ; y
      push  edx ; x
      invoke  ClientToScreen, thisPROC_hWnd[4 + SIZEOF RECT + SIZEOF POINT], esp
      ; (x, y) already on stack
      push  esp
      add DWORD PTR [esp], SIZEOF POINT
      call  PtInRect

      test  eax, eax
      lea esp, [esp + SIZEOF RECT]
      jne _x
      invoke  GetProp, thisPROC_hWnd[4],addr PROP_ORIGINAL_FONT
      invoke  SendMessage, thisPROC_hWnd[12], WM_SETFONT, eax, FALSE
      invoke  InvalidateRect, thisPROC_hWnd[8], NULL, FALSE
      invoke  ReleaseCapture
      _x: jmp CallWindowProc

      _start_capture:
      invoke  GetProp, thisPROC_hWnd[4],addr PROP_UNDERLINE_FONT
      invoke  SendMessage, thisPROC_hWnd[12], WM_SETFONT, eax, FALSE
      invoke  InvalidateRect, thisPROC_hWnd[8], NULL, FALSE
      invoke  SetCapture, thisPROC_hWnd
      jmp CallWindowProc


      __WM_SETCURSOR:
      ; Since IDC_HAND is not available on all operating systems,
      ; we will load the arrow cursor if IDC_HAND is not present.
      mov edx, IDC_HAND
      @@: invoke  LoadCursor, NULL, edx
      test  eax, eax
      mov edx, IDC_ARROW
      je  @B
      invoke  SetCursor, eax
      mov eax, TRUE
      ret 5*4
_HyperlinkProc                  endp
OPTION PROLOGUE:PROLOGUEDEF
OPTION EPILOGUE:EPILOGUEDEF
ConvertStaticToHyperlink        proc hwndParent:DWORD, uiCtlId:DWORD

      Invoke GetDlgItem, hwndParent, uiCtlId
      mov uiCtlId, eax ; need handle, not Id

      ; Subclass the parent so we can color the controls as we desire
      cmp hwndParent, NULL
      je  @F
          invoke  GetWindowLong, hwndParent, GWL_WNDPROC
      ;
      cmp eax, OFFSET _HyperlinkParentProc
      je  @F
          Invoke  SetProp, hwndParent,addr PROP_ORIGINAL_PROC, eax
          Invoke  SetWindowLong, hwndParent, GWL_WNDPROC, OFFSET _HyperlinkParentProc
      @@:
      ; Make sure the control will send notifications
      ;
      invoke  GetWindowLong, uiCtlId, GWL_STYLE
      or  eax, SS_NOTIFY
      invoke  SetWindowLong, uiCtlId, GWL_STYLE, eax

      ; Subclass the existing control

      invoke  GetWindowLong, uiCtlId, GWL_WNDPROC
      invoke  SetProp, uiCtlId,addr PROP_ORIGINAL_PROC, eax
      invoke  SetWindowLong, uiCtlId, GWL_WNDPROC, OFFSET _HyperlinkProc

      ; Create an updated font by adding an underline

      invoke  SendMessage, uiCtlId, WM_GETFONT, 0, 0
      push  eax
      invoke  SetProp, uiCtlId,addr PROP_ORIGINAL_FONT, eax
      pop edx
      sub esp, SIZEOF LOGFONT
      invoke  GetObject, edx, SIZEOF LOGFONT, esp
      mov [esp].LOGFONT.lfUnderline, TRUE

      invoke  CreateFontIndirect, esp
      invoke  SetProp, uiCtlId,addr PROP_UNDERLINE_FONT, eax
      add esp, SIZEOF LOGFONT

      mov eax, TRUE
      ret
ConvertStaticToHyperlink        endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;IniManager Procedure
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
ReadIniFile                     proc hTargetWindow:DWORD
                                ; Returns TRUE ON Success
;                                szSecPrevious               db "PREVIOUS",0
;                                szSecMainOptions            db "OPTIONS",0
;                                ;
;                                szKeyCopySource             db "CopySource",0      ;IDC_EDIT_PRE_SOURCE
;                                szKeyCopyTarget             db "CopyTarget",0      ;IDC_EDIT_PRE_TARGET
;                                szKeyCopySubFolders         db "CopySubFolders",0  ;IDC_BUTTON_PRE_SOURCE_FILEDLG
;                                szKeyCopyOverWrite          db "CopyOverwrite",0   ;IDC_BUTTON_PRE_SOURCE_OVERWR
;                                szKeyProgram                db "Program",0         ;IDC_EDIT_PRE_EXECUTE
;                                szKeyParameters             db "Parameters",0      ;IDC_EDIT_PRE_PARAMS
;                                szKeyParameters             db "WorkingDir",0      ;IDC_EDIT_PRE_WRKDIR
;                                szkeyShow                   db "Show",0            ;IDC_CHECK_PRE_SHOW
;                                szKeyCopyAfterExecute       db "CopyAfterExecute",0;IDC_CHECK_PRE_COPY_AFTER
;                                szkeyTimeout                db "Timeout",0         ;IDC_EDIT_PRE_TIMEOUT
;                                szkeyUnhide                 db "Unhide",0          ;IDC_CHECK_PRE_UNHIDE
;                                szkeyTerm                   db "Term",0            ;IDC_CHECK_PRE_TERM
;                                szkeyAskCopy                db "AskCopy",0         ;IDC_CHECK_PRE_ASK_COPY_AFTER
;                                szkeyAskExecute             db "AskExecute",0      ;IDC_CHECK_PRE_ASK_BEFORE_EXEC
;                                szKeyGuiOnCancelCopy        db "GuiOnCancelCopy",0 ;IDC_CHECK_PRE_GUI_COPY_AFTER
;                                szKeyGuiOnCancelExecute     db "GuiOnCancelExecute",0 ;IDC_CHECK_PRE_GUI_CANCEL_EXEC

        ; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
        ;                       PREVIUS TASK
        ; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
        ;******************** PRE CopySource Path **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyCopySource,0,addr szPreCopySource,sizeof szPreCopySource,addr szIniFile
        .if hTargetWindow != NULL
            Invoke SetDlgItemText,hTargetWindow,IDC_EDIT_PRE_SOURCE,addr szPreCopySource
        .endif
        ;
        ;******************** PRE CopySource SubFolders Valid values: TRUE or 1**************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyCopySubFolders,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2 ;2   The string pointed to by lpString1 is equal in lexical value to the string pointed to by lpString2.
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyCopySubFolders,1,addr szIniFile
        .endif
        mov dPreCopySubFolders,eax
        .if hTargetWindow != NULL
            .if dPreCopySubFolders
                Invoke CheckDlgButton,hTargetWindow,IDC_BUTTON_PRE_SOURCE_OVERWR,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_BUTTON_PRE_SOURCE_OVERWR,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE CopyTarget Path**************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyCopyTarget,0,addr szPreCopyTarget,sizeof szPreCopyTarget,addr szIniFile
        .if hTargetWindow != NULL
            Invoke SetDlgItemText,hTargetWindow,IDC_EDIT_PRE_TARGET,addr szPreCopyTarget
        .endif
        ;
        ;******************** PRE CopySource OverWrite **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyCopyOverWrite,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2 ;2   The string pointed to by lpString1 is equal in lexical value to the string pointed to by lpString2.
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyCopyOverWrite,1,addr szIniFile
        .endif
        mov dPreCopyOverwrite,eax
        .if hTargetWindow != NULL
            .if dPreCopyOverwrite
                Invoke CheckDlgButton,hTargetWindow,IDC_BUTTON_PRE_TARGET_OVERWR,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_BUTTON_PRE_TARGET_OVERWR,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE Run Program **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyProgram,0,addr szPreCmdRunProgram,sizeof szPreCmdRunProgram,addr szIniFile
        .if hTargetWindow != NULL
            Invoke SetDlgItemText,hTargetWindow,IDC_EDIT_PRE_EXECUTE,addr szPreCmdRunProgram
        .endif
        ;
        ;******************** PRE Run Parameters **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyParameters,0,addr szPreCmdRunParameters,sizeof szPreCmdRunParameters,addr szIniFile
        .if hTargetWindow != NULL
            Invoke SetDlgItemText,hTargetWindow,IDC_EDIT_PRE_PARAMS,addr szPreCmdRunParameters
        .endif
        ;
        ;******************** PRE Run Working Dir **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyWorkingDir,0,addr szPreCmdRunWorkingDir,sizeof szPreCmdRunWorkingDir,addr szIniFile
        .if hTargetWindow != NULL
            Invoke SetDlgItemText,hTargetWindow,IDC_EDIT_PRE_WRKDIR,addr szPreCmdRunWorkingDir
        .endif
        ;
        ;******************** PRE Unhide **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyUnhide,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyUnhide,0,addr szIniFile
        .endif
        mov dPreUnhide,eax
        .if hTargetWindow != NULL
            .if dPreUnhide
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_UNHIDE,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_UNHIDE,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE Show **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyShow,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            mov eax,SW_SHOW
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyShow,0,addr szIniFile
        .endif
        mov dPreShow,eax ;SW_SHOW
        .if hTargetWindow != NULL
            .if dPreShow
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_SHOW,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_SHOW,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE Ask Execute **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyAskExecute,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyAskExecute,0,addr szIniFile
        .endif
        mov dPreAskExecute,eax
        .if hTargetWindow != NULL
            .if dPreAskExecute
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_ASK_BEFORE_EXEC,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_ASK_BEFORE_EXEC,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE Gui On Cancel Execute **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyGuiOnCancelExecute ,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyGuiOnCancelExecute ,0,addr szIniFile
        .endif
        mov dPreGuiOnCancelExecute,eax
        .if hTargetWindow != NULL
            .if dPreGuiOnCancelExecute
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_GUI_CANCEL_EXEC,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_GUI_CANCEL_EXEC,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE Terminate **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyTerm,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyTerm,0,addr szIniFile
        .endif
        mov dPreTerm,eax
        .if hTargetWindow != NULL
            .if dPreTerm
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_TERM,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_TERM,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE Copy After Execute **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyCopyAfterExecute,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyCopyAfterExecute,0,addr szIniFile
        .endif
        mov dPreCopyAfterExecute,eax
        .if hTargetWindow != NULL
            .if dPreCopyAfterExecute
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_COPY_AFTER,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_COPY_AFTER,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE Ask Copy After **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyAskCopyAfter,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyAskCopyAfter,0,addr szIniFile
        .endif
        mov dPreAskCopyAfterExecute,eax
        .if hTargetWindow != NULL
            .if dPreAskCopyAfterExecute
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_ASK_COPY_AFTER,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_ASK_COPY_AFTER,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE GUI on Cancel Copy After **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyGuiOnCancelCopyAfter,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyGuiOnCancelCopyAfter,0,addr szIniFile
        .endif
        mov dPreGuiOnCancelCopyAfter,eax
        .if hTargetWindow != NULL
            .if dPreGuiOnCancelCopyAfter
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_GUI_COPY_AFTER,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_GUI_COPY_AFTER,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE TimeOut **************
        Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyTimeout,0,addr szIniFile
        mov dPreTimeout,eax
        .if hTargetWindow != NULL
            Invoke SetDlgItemInt,hTargetWindow,IDC_EDIT_PRE_TIMEOUT,dPreTimeout,FALSE ;TRUE=Singed, FALSE=Unsigned
        .endif
        ;
        ;
        ;******************** PRE Copy Before Execute **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyCopyBeforeExecute,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyCopyBeforeExecute,0,addr szIniFile
        .endif
        mov dPreCopyBeforeExecute,eax
        .if hTargetWindow != NULL
            .if dPreCopyBeforeExecute
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_COPY_BEFORE,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_COPY_BEFORE,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE Ask Copy Before **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyAskCopyBefore,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyAskCopyBefore,0,addr szIniFile
        .endif
        mov dPreAskCopyBeforeExecute,eax
        .if hTargetWindow != NULL
            .if dPreAskCopyBeforeExecute
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_ASK_COPY_BEFORE,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_ASK_COPY_BEFORE,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE GUI on Cancel Copy Before **************
        Invoke GetPrivateProfileString,addr szSecPrevious,addr szKeyGuiOnCancelCopyBefore,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecPrevious,addr szKeyGuiOnCancelCopyBefore,0,addr szIniFile
        .endif
        mov dPreGuiOnCancelCopyBefore,eax
        .if hTargetWindow != NULL
            .if dPreGuiOnCancelCopyBefore
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_GUI_COPY_BEFORE,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_PRE_GUI_COPY_BEFORE,BST_UNCHECKED
            .endif
        .endif
        ;
        ;
        ; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
        ;                       MAIN TASK
        ; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
        ;******************** CopySource Path **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyCopySource,0,addr szCopySource,sizeof szCopySource,addr szIniFile
        .if hTargetWindow != NULL
            Invoke SetDlgItemText,hTargetWindow,IDC_EDIT_MAIN_SOURCE,addr szCopySource
        .endif
        ;
        ;******************** CopySource SubFolders Valid values: TRUE or 1**************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyCopySubFolders,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2 ;2   The string pointed to by lpString1 is equal in lexical value to the string pointed to by lpString2.
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyCopySubFolders,1,addr szIniFile
        .endif
        mov dCopySubFolders,eax
        .if hTargetWindow != NULL
            .if dCopySubFolders
                Invoke CheckDlgButton,hTargetWindow,IDC_BUTTON_MAIN_SOURCE_OVERWR,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_BUTTON_MAIN_SOURCE_OVERWR,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** CopyTarget Path**************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyCopyTarget,0,addr szCopyTarget,sizeof szCopyTarget,addr szIniFile
        .if hTargetWindow != NULL
            Invoke SetDlgItemText,hTargetWindow,IDC_EDIT_MAIN_TARGET,addr szCopyTarget
        .endif
        ;
        ;******************** CopySource OverWrite **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyCopyOverWrite,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2 ;2   The string pointed to by lpString1 is equal in lexical value to the string pointed to by lpString2.
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyCopyOverWrite,1,addr szIniFile
        .endif
        mov dCopyOverwrite,eax
        .if hTargetWindow != NULL
            .if dCopyOverwrite
                Invoke CheckDlgButton,hTargetWindow,IDC_BUTTON_MAIN_TARGET_OVERWR,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_BUTTON_MAIN_TARGET_OVERWR,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** Run Program **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyProgram,0,addr szCmdRunProgram,sizeof szCmdRunProgram,addr szIniFile
        .if hTargetWindow != NULL
            Invoke SetDlgItemText,hTargetWindow,IDC_EDIT_MAIN_EXECUTE,addr szCmdRunProgram
        .endif
        ;
        ;******************** Run Parameters **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyParameters,0,addr szCmdRunParameters,sizeof szCmdRunParameters,addr szIniFile
        .if hTargetWindow != NULL
            Invoke SetDlgItemText,hTargetWindow,IDC_EDIT_MAIN_PARAMS,addr szCmdRunParameters
        .endif
        ;
        ;******************** Run Working Dir **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyWorkingDir,0,addr szCmdRunWorkingDir,sizeof szCmdRunWorkingDir,addr szIniFile
        .if hTargetWindow != NULL
            Invoke SetDlgItemText,hTargetWindow,IDC_EDIT_MAIN_WRKDIR,addr szCmdRunWorkingDir
        .endif
        ;
        ;******************** Unhide **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyUnhide,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyUnhide,0,addr szIniFile
        .endif
        mov dUnhide,eax
        .if hTargetWindow != NULL
            .if dUnhide
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_UNHIDE,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_UNHIDE,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** Show **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyShow,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            mov eax,SW_SHOW
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyShow,0,addr szIniFile
        .endif
        mov dShow,eax ;SW_SHOW
        .if hTargetWindow != NULL
            .if dShow
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_SHOW,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_SHOW,BST_UNCHECKED
            .endif
        .endif
        ;
        ;
        ;******************** Ask Execute **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyAskExecute,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyAskExecute,0,addr szIniFile
        .endif
        mov dAskExecute,eax
        .if hTargetWindow != NULL
            .if dAskExecute
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_ASK_BEFORE_EXEC,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_ASK_BEFORE_EXEC,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** Gui On Cancel Execute **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyGuiOnCancelExecute ,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyGuiOnCancelExecute ,0,addr szIniFile
        .endif
        mov dGuiOnCancelExecute,eax
        .if hTargetWindow != NULL
            .if dGuiOnCancelExecute
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_GUI_CANCEL_EXEC,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_GUI_CANCEL_EXEC,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** Terminate **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyTerm,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyTerm,0,addr szIniFile
        .endif
        mov dTerm,eax
        .if hTargetWindow != NULL
            .if dTerm
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_TERM,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_TERM,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** CopyAfterExecute **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyCopyAfterExecute,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyCopyAfterExecute,0,addr szIniFile
        .endif
        mov dCopyAfterExecute,eax
        .if hTargetWindow != NULL
            .if dCopyAfterExecute
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_COPY_AFTER,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_COPY_AFTER,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE Ask Copy After **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyAskCopyAfter,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyAskCopyAfter,0,addr szIniFile
        .endif
        mov dAskCopyAfterExecute,eax
        .if hTargetWindow != NULL
            .if dAskCopyAfterExecute
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_ASK_COPY_AFTER,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_ASK_COPY_AFTER,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** GUI on Cancel Copy After **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyGuiOnCancelCopyAfter,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyGuiOnCancelCopyAfter,0,addr szIniFile
        .endif
        mov dGuiOnCancelCopyAfter,eax
        .if hTargetWindow != NULL
            .if dGuiOnCancelCopyAfter
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_GUI_COPY_AFTER,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_GUI_COPY_AFTER,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE TimeOut **************
        Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyTimeout,0,addr szIniFile
        mov dTimeout,eax
        .if hTargetWindow != NULL
            Invoke SetDlgItemInt,hTargetWindow,IDC_EDIT_MAIN_TIMEOUT,dTimeout,FALSE ;TRUE=Singed, FALSE=Unsigned
        .endif
        ;
        ;******************** CopyBeforeExecute **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyCopyBeforeExecute,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyCopyBeforeExecute,0,addr szIniFile
        .endif
        mov dCopyBeforeExecute,eax
        .if hTargetWindow != NULL
            .if dCopyBeforeExecute
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_COPY_BEFORE,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_COPY_BEFORE,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** PRE Ask Copy Before **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyAskCopyBefore,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyAskCopyBefore,0,addr szIniFile
        .endif
        mov dAskCopyBeforeExecute,eax
        .if hTargetWindow != NULL
            .if dAskCopyBeforeExecute
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_ASK_COPY_BEFORE,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_ASK_COPY_BEFORE,BST_UNCHECKED
            .endif
        .endif
        ;
        ;******************** GUI on Cancel Copy Before **************
        Invoke GetPrivateProfileString,addr szSecMainOptions,addr szKeyGuiOnCancelCopyBefore,0,addr szCmdCompare,sizeof szCmdCompare,addr szIniFile
        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE, addr szTrue, -1,addr szCmdCompare, -1
        .if eax==2
            sub eax,1
        .else
            Invoke GetPrivateProfileInt,addr szSecMainOptions,addr szKeyGuiOnCancelCopyBefore,0,addr szIniFile
        .endif
        mov dGuiOnCancelCopyBefore,eax
        .if hTargetWindow != NULL
            .if dGuiOnCancelCopyBefore
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_GUI_COPY_BEFORE,BST_CHECKED
            .else
                Invoke CheckDlgButton,hTargetWindow,IDC_CHECK_MAIN_GUI_COPY_BEFORE,BST_UNCHECKED
            .endif
        .endif
        ;
        xor eax,eax
        add eax,1
        ExitReadIniFile:
    ret
ReadIniFile                     endp
align 4
WriteIniFile                    proc hTargetWindow:DWORD
                                ;
                                ;
        ; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
        ;                       PREVIUS TASK
        ; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
        ;
        ;******************** PRE CopySource Path **************
        Invoke GetDlgItemText,hTargetWindow,IDC_EDIT_PRE_SOURCE,addr szPreCopySource,sizeof szPreCopySource
        Invoke ValidateIniValue,addr szPreCopySource,sizeof szPreCopySource
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyCopySource,addr szPreCopySource,addr szIniFile
        .if eax==0 ;Returns 0 if fails.
            jmp ExitWriteIniFile
        .endif
        ;
        xor eax,eax
        mov szCmdCompare[1],al
        ;
        ;******************** PRE SubFolders Check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_BUTTON_PRE_SOURCE_OVERWR
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyCopySubFolders,addr szCmdCompare,addr szIniFile
        ;
        ;******************** PRE CopyTarget Check**************
        Invoke GetDlgItemText,hTargetWindow,IDC_EDIT_PRE_TARGET,addr szPreCopyTarget,sizeof szPreCopyTarget
        Invoke ValidateIniValue,addr szPreCopyTarget,sizeof szPreCopyTarget
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyCopyTarget,addr szPreCopyTarget,addr szIniFile
        ;
        ;
        ;******************** PRE OverWrite Check**************4
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_BUTTON_PRE_TARGET_OVERWR
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyCopyOverWrite,addr szCmdCompare,addr szIniFile
        ;
        ;******************** PRE RunProgram **************
        Invoke GetDlgItemText,hTargetWindow,IDC_EDIT_PRE_EXECUTE,addr szPreCmdRunProgram,sizeof szPreCmdRunProgram
        Invoke ValidateIniValue,addr szPreCmdRunProgram,sizeof szPreCmdRunProgram
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyProgram,addr szPreCmdRunProgram,addr szIniFile
        ;
        ;******************** PRE RunParameters **************
        Invoke GetDlgItemText,hTargetWindow,IDC_EDIT_PRE_PARAMS,addr szPreCmdRunParameters,sizeof szPreCmdRunParameters
        Invoke ValidateIniValue,addr szPreCmdRunParameters,sizeof szPreCmdRunParameters
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyParameters,addr szPreCmdRunParameters,addr szIniFile
        ;
        ;******************** PRE RunWorkingDir **************
        Invoke GetDlgItemText,hTargetWindow,IDC_EDIT_PRE_WRKDIR,addr szPreCmdRunWorkingDir,sizeof szPreCmdRunWorkingDir
        Invoke ValidateIniValue,addr szPreCmdRunWorkingDir,sizeof szPreCmdRunWorkingDir
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyWorkingDir,addr szPreCmdRunWorkingDir,addr szIniFile
        ;
        ;******************** PRE Unhide check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_PRE_UNHIDE
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyUnhide,addr szCmdCompare,addr szIniFile
        ;
        ;******************** PRE Show check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_PRE_SHOW
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyShow,addr szCmdCompare,addr szIniFile
        ;
        ;******************** PRE AskExecute check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_PRE_ASK_BEFORE_EXEC
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyAskExecute,addr szCmdCompare,addr szIniFile
        ;
        ;******************** PRE GUI on Cancel Execute check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_PRE_GUI_CANCEL_EXEC
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyGuiOnCancelExecute,addr szCmdCompare,addr szIniFile
        ;
        ;******************** PRE Term check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_PRE_TERM
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyTerm,addr szCmdCompare,addr szIniFile
        ;
        ;******************** PRE CopyAfterExecute check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_PRE_COPY_AFTER
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyCopyAfterExecute,addr szCmdCompare,addr szIniFile
        ;
        ;******************** PRE Ask Copy after check  **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_PRE_ASK_COPY_AFTER
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyAskCopyAfter,addr szCmdCompare,addr szIniFile
        ;
        ;******************** PRE Gui On Cancel Copy After  check**************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_PRE_GUI_COPY_AFTER
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyGuiOnCancelCopyAfter,addr szCmdCompare,addr szIniFile
        ;
        ;******************** PRE TimeOut **************
        Invoke GetDlgItemText,hTargetWindow,IDC_EDIT_PRE_TIMEOUT,addr szCmdCompare,sizeof szCmdCompare
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyTimeout,addr szCmdCompare,addr szIniFile
        xor eax,eax
        mov szCmdCompare[1],al
        ;
        ;******************** PRE CopyBeforeExecute check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_PRE_COPY_BEFORE
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyCopyBeforeExecute,addr szCmdCompare,addr szIniFile
        ;
        ;******************** PRE Ask Copy Before check  **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_PRE_ASK_COPY_BEFORE
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyAskCopyBefore,addr szCmdCompare,addr szIniFile
        ;
        ;******************** PRE Gui On Cancel Copy Before check**************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_PRE_GUI_COPY_BEFORE
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecPrevious,addr szKeyGuiOnCancelCopyBefore,addr szCmdCompare,addr szIniFile
        ;
        ;
        ; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
        ;                       MAIN TASK
        ; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
        ;
        ;******************** CopySource Path **************
        Invoke GetDlgItemText,hTargetWindow,IDC_EDIT_MAIN_SOURCE,addr szCopySource,sizeof szCopySource
        Invoke ValidateIniValue,addr szCopySource,sizeof szCopySource
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyCopySource,addr szCopySource,addr szIniFile
        .if eax==0 ;Returns 0 if fails.
            jmp ExitWriteIniFile
        .endif
        ;
        xor eax,eax
        mov szCmdCompare[1],al
        ;
        ;******************** SubFolders Check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_BUTTON_MAIN_SOURCE_OVERWR
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyCopySubFolders,addr szCmdCompare,addr szIniFile
        ;
        ;******************** CopyTarget Check**************
        Invoke GetDlgItemText,hTargetWindow,IDC_EDIT_MAIN_TARGET,addr szCopyTarget,sizeof szCopyTarget
        Invoke ValidateIniValue,addr szCopyTarget,sizeof szCopyTarget
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyCopyTarget,addr szCopyTarget,addr szIniFile
        ;
        ;
        ;******************** OverWrite Check**************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_BUTTON_MAIN_TARGET_OVERWR
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyCopyOverWrite,addr szCmdCompare,addr szIniFile
        ;
        ;******************** RunProgram **************
        Invoke GetDlgItemText,hTargetWindow,IDC_EDIT_MAIN_EXECUTE,addr szCmdRunProgram,sizeof szCmdRunProgram
        Invoke ValidateIniValue,addr szCmdRunProgram,sizeof szCmdRunProgram
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyProgram,addr szCmdRunProgram,addr szIniFile
        ;
        ;******************** RunParameters **************
        Invoke GetDlgItemText,hTargetWindow,IDC_EDIT_MAIN_PARAMS,addr szCmdRunParameters,sizeof szCmdRunParameters
        Invoke ValidateIniValue,addr szCmdRunParameters,sizeof szCmdRunParameters
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyParameters,addr szCmdRunParameters,addr szIniFile
        ;
        ;******************** RunWorkingDir **************
        Invoke GetDlgItemText,hTargetWindow,IDC_EDIT_MAIN_WRKDIR,addr szCmdRunWorkingDir,sizeof szCmdRunWorkingDir
        Invoke ValidateIniValue,addr szCmdRunWorkingDir,sizeof szCmdRunWorkingDir
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyWorkingDir,addr szCmdRunWorkingDir,addr szIniFile
        ;
        ;******************** Unhide check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_MAIN_UNHIDE
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyUnhide,addr szCmdCompare,addr szIniFile
        ;
        ;******************** Show check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_MAIN_SHOW
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyShow,addr szCmdCompare,addr szIniFile
        ;
        ;******************** Ask Execute check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_MAIN_ASK_BEFORE_EXEC
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyAskExecute,addr szCmdCompare,addr szIniFile
        ;
        ;******************** GUI on Cancel Ask Execute check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_MAIN_GUI_CANCEL_EXEC
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyGuiOnCancelExecute,addr szCmdCompare,addr szIniFile
        ;
        ;******************** Term check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_MAIN_TERM
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyTerm,addr szCmdCompare,addr szIniFile
        ;
        ;******************** CopyAfterExecute check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_MAIN_COPY_AFTER
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyCopyAfterExecute,addr szCmdCompare,addr szIniFile
        ;
        ;******************** AskCopy after check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_MAIN_ASK_COPY_AFTER
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyAskCopyAfter,addr szCmdCompare,addr szIniFile
        ;
        ;******************** Gui On Cancel Copy After check**************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_MAIN_GUI_COPY_AFTER
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyGuiOnCancelCopyAfter,addr szCmdCompare,addr szIniFile
        ;
        ;******************** TimeOut **************
        Invoke GetDlgItemText,hTargetWindow,IDC_EDIT_MAIN_TIMEOUT,addr szCmdCompare,sizeof szCmdCompare
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyTimeout,addr szCmdCompare,addr szIniFile
        xor eax,eax
        mov szCmdCompare[1],al
        ;
        ;******************** CopyBeforeExecute check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_MAIN_COPY_BEFORE
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyCopyBeforeExecute,addr szCmdCompare,addr szIniFile
        ;
        ;******************** Ask Copy Before check **************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_MAIN_ASK_COPY_BEFORE
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyAskCopyBefore,addr szCmdCompare,addr szIniFile
        ;
        ;******************** Gui On Cancel Copy Before check**************
        Invoke IsDlgButtonChecked,hTargetWindow,IDC_CHECK_MAIN_GUI_COPY_BEFORE
        add al,'0'
        mov szCmdCompare[0],al
        Invoke WritePrivateProfileString,addr szSecMainOptions,addr szKeyGuiOnCancelCopyBefore,addr szCmdCompare,addr szIniFile
        ;
        ;
        xor eax,eax
        add eax,1
        ExitWriteIniFile:
    ret
WriteIniFile                    endp
align 4
ValidateIniValue                proc pSource:DWORD,dSize:DWORD
                                ;
                                LOCAL       szValidBuffer[4096]:BYTE
                                LOCAL       bCopyBack:DWORD
                                LOCAL       bSameSize:DWORD
                                LOCAL       bAddExtra:DWORD
                                szText      szQuoteSimple,"'"
                                szText      szQuoteDoble,'"'
                                ;Converts "C:\MyApp\" to ""C:\MyApp\"" or
                                ;Converts 'My quote text' to ''My quote text''
                                ;so when stored in ini file can be retrieve back to single quotes.
                                ;https://en.wikipedia.org/wiki/INI_file
                                ;
        Invoke FillBuffer,addr szValidBuffer,sizeof szValidBuffer,0
        mov bSameSize,FALSE
        mov bCopyBack,FALSE
        mov bAddExtra,FALSE
        invoke lstrlen,pSource
        .if eax==0
            jmp ExitValidateIniValue
        .elseif eax > 4095 ;Larger than buffer?
            mov eax,-1
            jmp ExitValidateIniValue
        .endif

        invoke lstrlen,pSource
        mov ecx,eax
        mov edx,pSource
        .if byte ptr[edx] == '"'
            .if byte ptr[edx+ecx-1] == '"'
                invoke lstrcpy,addr szValidBuffer,addr szQuoteDoble
                invoke lstrcat,addr szValidBuffer,pSource
                invoke lstrcat,addr szValidBuffer,addr szQuoteDoble
                invoke lstrcpy,pSource,addr szValidBuffer
            .else
                ;invoke lstrcpy,addr szValidBuffer,addr szQuoteDoble
                ;invoke lstrcat,addr szValidBuffer,pSource
                ;invoke lstrcpy,pSource,addr szValidBuffer
            .endif
        .elseif byte ptr[edx+ecx-1] == '"'
            ;invoke lstrcpy,addr szValidBuffer,pSource
            ;invoke lstrcat,addr szValidBuffer,addr szQuoteDoble
            ;invoke lstrcpy,pSource,addr szValidBuffer
        .endif
        invoke lstrlen,pSource
        mov ecx,eax
        mov edx,pSource
        .if byte ptr[edx] == "'"
            .if byte ptr[edx+ecx-1] == "'"
                invoke lstrcpy,addr szValidBuffer,addr szQuoteSimple
                invoke lstrcat,addr szValidBuffer,pSource
                invoke lstrcat,addr szValidBuffer,addr szQuoteSimple
                invoke lstrcpy,pSource,addr szValidBuffer
            .else
                ;invoke lstrcpy,addr szValidBuffer,addr szQuoteDoble
                ;invoke lstrcat,addr szValidBuffer,pSource
                ;invoke lstrcpy,pSource,addr szValidBuffer
            .endif
        .elseif byte ptr[edx+ecx-1] == '"'
            ;invoke lstrcpy,addr szValidBuffer,pSource
            ;invoke lstrcat,addr szValidBuffer,addr szQuoteDoble
            ;invoke lstrcpy,pSource,addr szValidBuffer
        .endif
        jmp ExitValidateIniValue
        ExitValidateIniValue:
        ;Invoke OutputDebugString,pSource
        ;Return 0 on success or Required string size
    ret
ValidateIniValue                endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
;           *** SolveMacros **
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
SolveMacros                     proc lpSourceStr:DWORD,lpTargetStr:DWORD
                                ;
                                LOCAL     szDummyBuffer[MAX_PATH]:BYTE
                                LOCAL     szTargetBuffer[MAX_PATH]:BYTE
                                LOCAL     dSize :DWORD
                                LOCAL     hReg:DWORD
                                LOCAL     LocalTime:SYSTEMTIME
;
;                                szText    szCurrentVersion,"Software\Microsoft\Windows\CurrentVersion"
;                                szText    szProgramFilesDir,"ProgramFilesDir"
;                                szText    szTwoDigitFormat,"%02d"                 ; To format '1' into '01:'
        ;
        Invoke lstrcpy,addr szTargetBuffer,lpSourceStr ;Source to Temp
        Invoke GetLocalTime, addr LocalTime
        ;ммммммммммммммммммммммммммммм YEAR мммммммммммммммммммммммм
        xor eax,eax
        mov ax,LocalTime.wYear
        mov dSize,eax
        Invoke DwordToAsccii,dSize,addr szDummyBuffer
        Invoke ReplaceMacros,addr szTargetBuffer,addr szYEAR,addr szDummyBuffer          ;Current year
        ;ммммммммммммммммммммммммммммм MONTH мммммммммммммммммммммммм
        xor eax,eax
        mov ax,LocalTime.wMonth
        lea ecx,szDummyBuffer
        add ecx,4
        Invoke wsprintfA, ecx, addr szTwoDigitFormat,eax
        lea ecx,szDummyBuffer
        add ecx,4
        Invoke ReplaceMacros,addr szTargetBuffer,addr szMONTH,ecx                                   ;Month of the Year (1-12)
        ;ммммммммммммммммммммммммммммм DAY мммммммммммммммммммммммм
        xor eax,eax
        mov ax,LocalTime.wDay
        lea ecx,szDummyBuffer
        add ecx,6
        Invoke wsprintfA, ecx, addr szTwoDigitFormat,eax
        lea ecx,szDummyBuffer
        add ecx,6
        Invoke ReplaceMacros,addr szTargetBuffer,addr szDAY,ecx                                     ;Day of the month (1-31)
        ;ммммммммммммммммммммммммммммм DATE мммммммммммммммммммммммм
        Invoke ReplaceMacros,addr szTargetBuffer,addr szDATE,addr szDummyBuffer          ;Current date (in the format YYYYMMDD)
        ;ммммммммммммммммммммммммммммм HOUR мммммммммммммммммммммммм
        xor eax,eax
        mov ax,LocalTime.wHour
        Invoke wsprintfA,addr szDummyBuffer, addr szTwoDigitFormat,eax
        Invoke ReplaceMacros,addr szTargetBuffer,addr szHOUR,addr szDummyBuffer          ;LocalTime Hour 0-24
        ;Invoke MessageBox,NULL,lpSourceStr,addr szDummyBuffer,NULL
        ;ммммммммммммммммммммммммммммм MINUTE мммммммммммммммммммммммм
        xor eax,eax
        mov ax,LocalTime.wMinute
        lea ecx,szDummyBuffer
        add ecx,2
        Invoke wsprintfA, ecx, addr szTwoDigitFormat,eax
        lea ecx,szDummyBuffer
        add ecx,2
        Invoke ReplaceMacros,addr szTargetBuffer,addr szMINUTE,ecx                                      ;LocalTime Hour 0-59
        ;ммммммммммммммммммммммммммммм SECOND мммммммммммммммммммммммм
        xor eax,eax
        mov ax,LocalTime.wSecond
        lea ecx,szDummyBuffer
        add ecx,4
        Invoke wsprintfA, ecx, addr szTwoDigitFormat,eax
        lea ecx,szDummyBuffer
        add ecx,4
        Invoke ReplaceMacros,addr szTargetBuffer,addr szSECOND,ecx                                      ;LocalTime Second 0-59
        ;ммммммммммммммммммммммммммммм TIME мммммммммммммммммммммммм
        Invoke ReplaceMacros,addr szTargetBuffer,addr szTIME,addr szDummyBuffer        ;Current time (in the format HHMMSS)
        ;ммммммммммммммммммммммммммммм WEEKDAY мммммммммммммммммммммммм
        xor eax,eax
        mov ax,WORD ptr [LocalTime.wDayOfWeek]
        inc ax
        mov dSize,eax
        Invoke DwordToAsccii,dSize,addr szDummyBuffer
        Invoke ReplaceMacros,addr szTargetBuffer,addr szWEEKDAY,addr szDummyBuffer     ;Days since Sunday (1 Ц 7)
        ;ммммммммммммммммммммммммммммм PROCESSID мммммммммммммммммммммммм
        Invoke DwordToAsccii,hProcessID_ForMacro,addr szDummyBuffer
        Invoke ReplaceMacros,addr szTargetBuffer,addr szPROCESSID,addr szDummyBuffer   ;ProcessID
        ;ммммммммммммммммммммммммммммм DRIVE мммммммммммммммммммммммм
        ;Invoke ReplaceMacros,addr szTargetBuffer,addr szDRIVE,addr szMapToDrive       ;Drive that is redirected to ShareAddress (D  db  C  db  etc)
        ;ммммммммммммммммммммммммммммм ERRORCODE мммммммммммммммммммммммм
        Invoke DwordToAsccii, dErrorCode,addr szDummyBuffer
        Invoke ReplaceMacros,addr szTargetBuffer,addr szERRORCODE,addr szDummyBuffer   ;Last Error Code Returned by  GetLastError
        ;ммммммммммммммммммммммммммммм ERRORTEXT мммммммммммммммммммммммм
        Invoke FormatMessage, FORMAT_MESSAGE_FROM_SYSTEM, 0, dErrorCode, 0, addr szDummyBuffer,MAX_PATH, 0
        .if szDummyBuffer
              Invoke ReplaceMacros,addr szTargetBuffer,addr szERRORTEXT,addr szDummyBuffer   ;Last Error Text Returned by  GetLastError
        .endif
        ;ммммммммммммммммммммммммммммм MYDOCUMENTS мммммммммммммммммммммммм
        Invoke SHGetFolderPathA,NULL,CSIDL_MYDOCUMENTS,NULL,0,addr szDummyBuffer   ;SHGFP_TYPE_CURRENT = 0
        ;A pointer to a null-terminated string of length MAX_PATH which will receive the path. If an error occurs or S_FALSE is returned, this string will be empty.
        ;The returned path does not include a trailing backslash. For example, "C:\Users" is returned rather than "C:\Users\".
        .if eax==0 ;If this function succeeds, it returns S_OK=0. Otherwise, it returns an HRESULT error code.
            Invoke ReplaceMacros,addr szTargetBuffer,addr szMYDOCUMENTS,addr szDummyBuffer  ;Path of the Program Working directory
        .endif        
;        Invoke GetVersion
;        and eax,00FFH 
;        .if eax < 6 ; Pre Vista
;            Invoke SHGetFolderPathA,NULL,CSIDL_MYDOCUMENTS,NULL,0,addr szDummyBuffer   ;SHGFP_TYPE_CURRENT = 0
;            ;A pointer to a null-terminated string of length MAX_PATH which will receive the path. If an error occurs or S_FALSE is returned, this string will be empty.
;            ;The returned path does not include a trailing backslash. For example, "C:\Users" is returned rather than "C:\Users\".
;            .if eax==0 ;If this function succeeds, it returns S_OK=0. Otherwise, it returns an HRESULT error code.
;                Invoke ReplaceMacros,addr szTargetBuffer,addr szMYDOCUMENTS,addr szDummyBuffer  ;Path of the Program Working directory
;            .endif
;        .else ;Vista+
;             FOLDERID_Documents          GUID  <0FDD39AD0h,238Fh,46AFh,<0ADh,0B4h,06Ch,85h,48h,03h,69h,0C7h>>
;            Invoke SHGetKnownFolderPath,addr FOLDERID_Documents,0,NULL,addr Ppwstring    ;KF_FLAG_DEFAULT = 0
;            ;Ppwstring When this method returns, contains the address of a pointer to a null-terminated Unicode string that specifies the path of the known folder.
;            ;The calling process is responsible for freeing this resource once it is no longer needed by calling CoTaskMemFree, whether SHGetKnownFolderPath succeeds or not.
;            ;The returned path does not include a trailing backslash. For example, "C:\Users" is returned rather than "C:\Users\".            
;            .if eax==0;Returns S_OK if successful, or an error
;                Invoke ReplaceMacros,addr szTargetBuffer,addr szMYDOCUMENTS,addr szDummyBuffer  ;Path of the Program Working directory
;            .endif
;        .endif
        ;ммммммммммммммммммммммммммммм PROGRAMDIR мммммммммммммммммммммммм
        Invoke GetPathOnly,addr szProgramPath,addr szDummyBuffer
        Invoke lstrlen, addr szDummyBuffer
        lea ecx,szDummyBuffer
        dec eax                                           ;Decrement the string
        cmp byte ptr [ecx+eax],'\'
        jne @F
            mov byte ptr [ecx+eax],00H
        @@:
        Invoke ReplaceMacros,addr szTargetBuffer,addr szPROGRAMDIR,addr szDummyBuffer  ;Path of the Program Working directory
        ;ммммммммммммммммммммммммммммм PROGRAMFILES мммммммммммммммммммммммм
        Invoke RegOpenKeyEx,HKEY_LOCAL_MACHINE,addr szCurrentVersion,NULL,KEY_READ,addr hReg
        .if eax==ERROR_SUCCESS
            mov dSize,MAX_PATH
            Invoke RegQueryValueEx,hReg,addr szProgramFilesDir,NULL,NULL,addr szDummyBuffer,addr dSize
            Invoke RegCloseKey,hReg
            Invoke lstrlen, addr szDummyBuffer
            lea ecx,szDummyBuffer
            dec eax                                           ;Decrement the string
            cmp byte ptr [ecx+eax],'\'
            jne @F
                mov byte ptr [ecx+eax],00H
            @@:
            Invoke ReplaceMacros,addr szTargetBuffer,addr szPROGRAMFILES,addr szDummyBuffer;Path of the Program Files Directory
        .endif
        ;ммммммммммммммммммммммммммммм PROGRAMNAME мммммммммммммммммммммммм
        Invoke ReplaceMacros,addr szTargetBuffer,addr szPROGRAMNAME,addr szProgramName  ;Name of the current executable file  usually = MapDrive
        ;ммммммммммммммммммммммммммммм SHARE мммммммммммммммммммммммм
        ;Invoke ReplaceMacros,addr szTargetBuffer,addr szSHARE,addr szShareAddress       ;Full Share Address
        ;ммммммммммммммммммммммммммммм SYSTEMDIR мммммммммммммммммммммммм
        Invoke GetSystemDirectory,addr szDummyBuffer,MAX_PATH
        Invoke lstrlen, addr szDummyBuffer
        lea ecx,szDummyBuffer
        dec eax                                           ;Decrement the string
        cmp byte ptr [ecx+eax],'\'
        jne @F
            mov byte ptr [ecx+eax],00H
        @@:
        Invoke ReplaceMacros,addr szTargetBuffer,addr szSYSTEMDIR,addr szDummyBuffer     ;Path of the Windows system directory
        ;ммммммммммммммммммммммммммммм SYSTEMDRIVE мммммммммммммммммммммммм
        lea ecx,szDummyBuffer
        add ecx,3
        mov byte ptr [ecx+eax],00H
        Invoke ReplaceMacros,addr szTargetBuffer,addr szSYSTEMDRIVE,addr szDummyBuffer   ;Default Drive Directory usually = C:\
        ;ммммммммммммммммммммммммммммм TEMPDIR мммммммммммммммммммммммм
        Invoke GetTempPath,MAX_PATH,addr szDummyBuffer
        Invoke lstrlen, addr szDummyBuffer
        lea ecx,szDummyBuffer
        dec eax                                           ;Decrement the string
        cmp byte ptr [ecx+eax],'\'
        jne @F
            mov byte ptr [ecx+eax],00H
        @@:
        Invoke ReplaceMacros,addr szTargetBuffer,addr szTEMPDIR,addr szDummyBuffer       ;path of the directory designated for temporary files
        ;ммммммммммммммммммммммммммммм USERNAME мммммммммммммммммммммммм
        mov dSize,MAX_PATH
        Invoke GetUserName,addr szDummyBuffer,addr dSize
        Invoke ReplaceMacros,addr szTargetBuffer,addr szUSERNAME,addr szDummyBuffer      ;Current user's Windows user ID
        ;ммммммммммммммммммммммммммммм COMPUTERNAME мммммммммммммммммммм
        mov dSize,MAX_PATH
        Invoke GetComputerName,addr szDummyBuffer,addr dSize
        Invoke ReplaceMacros,addr szTargetBuffer,addr szCOMPUTERNAME,addr szDummyBuffer  ;Computer name of the current system
        ;ммммммммммммммммммммммммммммм WINDIR мммммммммммммммммммммммм
        Invoke GetWindowsDirectory,addr szWindowsDir,sizeof szWindowsDir
            Invoke lstrlen, addr szWindowsDir
            lea ecx,szWindowsDir
            dec eax                                           ;Decrement the string
            cmp byte ptr [ecx+eax],'\'
            jne @F
                ;inc eax
                ;mov byte ptr [ecx+eax],'\'
                ;inc eax
                mov byte ptr [ecx+eax],00H
            @@:
        Invoke ReplaceMacros,addr szTargetBuffer,addr szWINDIR ,addr szWindowsDir           ;Path of the Windows directory
        ;
        Invoke lstrcpy,lpTargetStr,addr szTargetBuffer
        ;
    ret
SolveMacros                     endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
;           *** ReplaceMacros **
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
ReplaceMacros                   proc  lpSourceStr:DWORD, lpPattern:DWORD, lpTarget:DWORD
                                ;
                                LOCAL pLen:DWORD
                                LOCAL dTextPos:DWORD
                                LOCAL szTempCopy[1204]:BYTE
                                LOCAL szDummyBuffer[1024]:BYTE
        ;
        Invoke lstrcmp,lpSourceStr,lpPattern
        .if eax==0 ;NOTHING to Solve!
            Invoke lstrcpy,lpSourceStr,lpTarget
            ret
        .endif
        Invoke StrLen,lpPattern
        mov pLen, eax                       ; Pattern length
        @@:
            Invoke lstrcpy,addr szTempCopy,lpSourceStr ;Source to Temp
            Invoke InString,1,addr szTempCopy,lpPattern  ;0=No found, -1=same length or longer -2=out of range, >0 = TextPos
            cmp eax,1                         ;Bail if less than 1
            jl  @F
                mov dTextPos,eax              ;it returns the 1 based index of the start of the substring.
                Invoke lstrcpyn, addr szDummyBuffer,addr szTempCopy,eax  ;Copy leading char on Buffer
                Invoke lstrcat,addr szDummyBuffer,lpTarget           ;Appends lpTarget on Buffer
                mov eax,dTextPos              ;First char position of Pattern
                dec eax                       ;Get 0 based index
                mov ecx,pLen
                add eax,ecx                   ;eax==Position of final Pattern char in lpSourceStr
                lea ecx,szTempCopy            ;Address of lpSourceStr
                add ecx,eax                   ;ecx==Address' Position of final Pattern char in lpSourceStr
                Invoke lstrcat,addr szDummyBuffer,ecx  ;Appends Rest of lpTarget on Buffer
                Invoke lstrcpy,lpSourceStr,addr szDummyBuffer ;Copy back Buffer to Target
                cmp eax,NULL ;fail???
                je @F
            jmp @B
        @@:
        ;Invoke MessageBox,NULL,lpPattern,addr szTempCopy,MB_OK OR MB_ICONEXCLAMATION
  ret
ReplaceMacros                   endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
;           *** PasteMacro **
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
PasteMacro                      proc  dPopUpSelection :DWORD, hEditHandle :DWORD, dEditID
                                ;
                                LOCAL wSelStartPos:WORD
                                LOCAL wSelEndPos:WORD
                                LOCAL pMacroText:DWORD
                                LOCAL szTempCopy[MAX_PATH]:BYTE
                                LOCAL szDummyBuffer[MAX_PATH]:BYTE
        ;
        .if     dPopUpSelection == IDM_COMPUTERNAME
            mov eax, OFFSET szCOMPUTERNAME  ;Computer name of the current system
        .elseif dPopUpSelection == IDM_DATE
            mov eax, OFFSET szDATE          ;Current date (in the format YYYYMMDD)
        .elseif dPopUpSelection == IDM_DAY
            mov eax, OFFSET szDAY           ;Day of the month (1-31)
        .elseif dPopUpSelection == IDM_DRIVE
            mov eax, OFFSET szDRIVE         ;Drive that is redirected to ShareAddress (D  db  C  db  etc)
        .elseif dPopUpSelection == IDM_ERRORCODE
            mov eax, OFFSET szERRORCODE     ;Last Error Code Returned by  GetLastError
        .elseif dPopUpSelection == IDM_ERRORTEXT
            mov eax, OFFSET szERRORTEXT     ;Last Error Text Returned by  GetLastError
        .elseif dPopUpSelection == IDM_HOUR
            mov eax, OFFSET szHOUR                ;LocalTime Hour 0-24
        .elseif dPopUpSelection == IDM_MINUTE
            mov eax, OFFSET szMINUTE          ;LocalTime Hour Minute 00-59
        .elseif dPopUpSelection == IDM_MONTH
            mov eax, OFFSET szMONTH         ;Month of the Year (1-12)
        .elseif dPopUpSelection ==  IDM_MYDOCUMENTS
            mov eax, OFFSET szMYDOCUMENTS             
        .elseif dPopUpSelection == IDM_PROCESSID
            mov eax, OFFSET szPROCESSID         ;Process ID
        .elseif dPopUpSelection == IDM_PROGRAMDIR
            mov eax, OFFSET szPROGRAMDIR    ;Path of the Program Working directory
        .elseif dPopUpSelection == IDM_PROGRAMFILES
            mov eax, OFFSET szPROGRAMFILES  ;Path of the Program Files Directory
        .elseif dPopUpSelection == IDM_PROGRAMNAME
            mov eax, OFFSET szPROGRAMNAME   ;Name of the current executable file  usually = MapDrive
        .elseif dPopUpSelection == IDM_SECOND
            mov eax, OFFSET szSECOND          ;LocalTime Second 00-59
        .elseif dPopUpSelection == IDM_SHARE
            mov eax, OFFSET szSHARE         ;Full Share Address
        .elseif dPopUpSelection == IDM_SYSTEMDIR
            mov eax, OFFSET szSYSTEMDIR     ;Path of the Windows system directory
        .elseif dPopUpSelection == IDM_SYSTEMDRIVE
            mov eax, OFFSET szSYSTEMDRIVE   ;Default Root Drive usually = C:\
        .elseif dPopUpSelection == IDM_TEMPDIR
            mov eax, OFFSET szTEMPDIR          ;Default Temp Directory
        .elseif dPopUpSelection == IDM_TIME
            mov eax, OFFSET szTIME          ;Current time (in the format HHMMSS)
        .elseif dPopUpSelection == IDM_USERNAME
            mov eax, OFFSET szUSERNAME      ;Current user's Windows user ID
        .elseif dPopUpSelection == IDM_WEEKDAY
            mov eax, OFFSET szWEEKDAY       ;Days since Sunday (1 Ц 7)
        .elseif dPopUpSelection == IDM_WINDIR
            mov eax, OFFSET szWINDIR        ;Path of the Windows directory
        .elseif dPopUpSelection == IDM_YEAR
            mov eax, OFFSET szYEAR          ;Current year
        .else
            ret
        .endif
        ;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        mov pMacroText,eax
        mov szTempCopy,0
        Invoke SendMessage,hEditHandle,EM_GETSEL,0,0
        mov wSelStartPos,ax
        shr eax,16
        mov wSelEndPos,ax
        Invoke GetDlgItemText,hOwnerPopMenu,dEditID,addr  szDummyBuffer,MAX_PATH ;Copy contenst of edit
        xor eax,eax
        mov ax,wSelStartPos
        inc eax
        Invoke GetDlgItemText,hOwnerPopMenu,dEditID,addr  szTempCopy,eax ;Get leading text before caret
        Invoke lstrcat,addr szTempCopy,pMacroText                                    ;Appends lpTarget on Buffer
        xor eax,eax
        mov ax,wSelEndPos
        lea ecx,szDummyBuffer
        add eax,ecx
        Invoke lstrcat,addr szTempCopy,eax                                   ;Appends rest of text
        Invoke SetDlgItemText,hOwnerPopMenu,dEditID,addr  szTempCopy
        Invoke StrLen,pMacroText
        xor ecx,ecx
        mov cx,wSelEndPos
        add eax,ecx
        Invoke SendMessage,hEditHandle,EM_SETSEL,eax,eax
    ;
    ret
PasteMacro                      endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;SubWndProc
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
SubWndProc                      proc hWnd:DWORD,uMsg:DWORD,wParam :DWORD,lParam :DWORD
                                ;le
                                LOCAL dCtrlID:DWORD
                                LOCAL   pt:POINT
                                LOCAL   rc:RECT
        ;
        .if uMsg==WM_CHAR
            ;
            .if wParam==VK_SPACE
                Invoke GetKeyState,VK_CONTROL
                and ax,0FFFEH
                .if ax
                    Invoke GetFocus
                    mov hFocus,eax
                    Invoke GetWindowRect,hFocus,addr rc
                    Invoke GetDlgCtrlID, hFocus
                    mov dCtrlID,eax
                    ;
                    Invoke GetCaretPos,addr pt
                    Invoke SetForegroundWindow, hOwnerPopMenu
                    mov eax,rc.left
                    mov ecx,pt.x
                    add eax,ecx
                    mov pt.x,eax
                    ;
                    mov eax,rc.bottom
                    mov ecx,rc.top
                    sub eax,ecx
                    mov ecx,rc.top
                    add eax,ecx
                    mov ecx,pt.y
                    add eax,ecx
                    mov pt.y,eax
                    ;
                    Invoke TrackPopupMenu, hPopupMenu, TPM_LEFTALIGN + TPM_RETURNCMD, pt.x, pt.y,NULL, hOwnerPopMenu, NULL
                    .if eax
                        Invoke PasteMacro,eax,hFocus,dCtrlID  ; PopUpSelection,EditHandle,EditID
                    .endif
                    xor eax,eax
                    ret
                .endif
            .endif
            ;
        .endif
        ;
        mov ecx,hWnd
        lea eax,ActiveIDC
        @@:
            cmp dword ptr[eax],ecx
            je @F
            add eax,8
            jmp @B
        @@:
        add eax,4
        mov eax,dword ptr[eax]
        Invoke CallWindowProc,eax,hWnd,uMsg,wParam,lParam
        ;
    ret
SubWndProc                      endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;GETFILENAME  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
FileDialog                      proc hParent:DWORD,pszTitle:DWORD,pszFileName:DWORD,dFileNameSize:DWORD,pszFilter:DWORD,dFilterSize:DWORD,dBrowseForFolder:DWORD
                                ; Returns File
                                LOCAL pIDList :DWORD
                                LOCAL bi:BROWSEINFO
                                ;
                                LOCAL ofn:OPENFILENAME

        ;
        .if dBrowseForFolder == TRUE
            ;
            Invoke FillBuffer,addr bi,sizeof BROWSEINFO,0
            ;
            mov eax,                hParent         ; parent handle
            mov bi.hwndOwner,       eax
            mov bi.pidlRoot,        0
            mov bi.pszDisplayName,  0
            mov eax,                pszFilter       ; secondary text
            mov bi.lpszTitle,       eax
            mov bi.ulFlags,         BIF_RETURNONLYFSDIRS ;OR BIF_DONTGOBELOWDOMAIN
            mov bi.lpfn,            offset cbBrowseWindow
            mov eax,                pszTitle        ; main title
            mov bi.lParam,          eax
            mov bi.iImage,          0
            ;
            invoke SHBrowseForFolder,addr bi
            mov pIDList, eax
            ;
            .if pIDList == 0
              mov eax, 0      ; if CANCEL return FALSE
              push eax
              jmp @F
            .else
              Invoke SHGetPathFromIDList,pIDList,pszFileName
              mov eax, 1        ; if OK, return TRUE
              push eax
              jmp @F
            .endif
            ;
            @@:
            ;
            Invoke CoTaskMemFree,pIDList
            ;
            pop eax
            .if eax!=0
                Invoke lstrlen,pszFileName
                mov ecx,pszFileName
                cmp byte ptr [ecx+eax-1],'\'
                je @F
                    .if eax < MAX_PATH
                        mov byte ptr [ecx+eax],'\'
                        mov byte ptr [ecx+eax+1],0
                    .endif
                @@:
                xor eax,eax
                add eax,1
            .endif
        .else
            Invoke FillBuffer,addr ofn, sizeof  OPENFILENAME,0
            mov ofn.lStructSize,        sizeof OPENFILENAME
            mov eax,hParent
            mov ofn.hWndOwner,eax
            Invoke GetModuleHandle,NULL
            mov ofn.hInstance,eax
            ;
            mov ecx,pszTitle
            mov ofn.lpstrTitle,ecx
            ;
            mov edx,pszFileName
            mov ofn.lpstrFile,edx
            ;
            mov eax,dFileNameSize
            mov ofn.nMaxFile,eax
            ;
            mov ecx,pszFilter
            mov ofn.lpstrFilter,ecx
            ;
            mov ofn.Flags,              OFN_EXPLORER OR OFN_LONGNAMES OR OFN_NOCHANGEDIR ;or OFN_FILEMUSTEXIST
            ;
            Invoke GetOpenFileNameA,addr ofn
        .endif
        ;
    ret
    ;
FileDialog                      endp
; #########################################################################
align 4
cbBrowseWindow                  proc hWin:DWORD,uMsg:DWORD,lParam:DWORD,lpData :DWORD
                                ;
        ;
        Invoke SetWindowText,hWin,lpData
    ret
cbBrowseWindow                  endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;CopyFiles
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
CopyFiles                       proc pszSourcePath:DWORD,pszTargetPath:DWORD,dSubFolders:DWORD,dOverWrite
                                ;
                                LOCAL bIsADirectory:BYTE
                                LOCAL szBaseWildcard[MAX_PATH]:BYTE
                                LOCAL szTrailWildcard[MAX_PATH]:BYTE
                                LOCAL szSingleWildcard[MAX_PATH]:BYTE
        ;
        mov bIsADirectory,1                  ;by default try to copy a directory ;0=File,1=Directory,2=WildCards
        Invoke lstrlen,pszSourcePath
        .if eax==0
            jmp ExitCopyFiles
        .endif
        ;
        dec eax
        mov ecx,pszSourcePath
        cmp byte ptr[ecx+eax],'\'
        jz IsDirectory
            @@:
                sub eax,1
                cmp byte ptr[ecx+eax],'*'
                mov bIsADirectory,2 ;0=File,1=Directory,2=WildCards
                jz IsDirectory
                cmp byte ptr[ecx+eax],'?'
                mov bIsADirectory,2 ;0=File,1=Directory,2=WildCards
                jz IsDirectory
                cmp byte ptr[ecx+eax],';'
                mov bIsADirectory,2 ;0=File,1=Directory,2=WildCards
                jz IsDirectory
                cmp byte ptr[ecx+eax],'\'
                jz @F
                cmp eax,0
                jz @F
                jmp @B
            @@:
            Invoke FileExist,pszSourcePath
            .if eax == 1 ; File exists
                mov bIsADirectory,0     ;For Sure is NOT a Directory or WildCard Path
            .else
                jmp ExitCopyFiles     ;File Don't Exist, or It's NOT a Wild Card
            .endif
        IsDirectory:        ;A this point the directory(Wildcard) / archive is valid
        .if bIsADirectory==0  ;0=File,1=Directory,2=WildCards
              Invoke SingleCopy,pszSourcePath,pszTargetPath,dOverWrite
        .else
            ;*************** Verify that the Target has the leading "\"
            Invoke lstrlen,pszTargetPath
            dec eax
            mov ecx,pszTargetPath
            cmp byte ptr [ecx+eax],05Ch ;'\'
            je  @F
                inc eax
                mov byte ptr[ecx+eax],05Ch
                inc eax
                mov byte ptr[ecx+eax],0
            @@:
            .if bIsADirectory==1
                Invoke lstrlen,pszSourcePath
                mov ecx,pszSourcePath
                mov byte ptr [eax+ecx],'*'
                inc eax
                mov byte ptr [eax+ecx],'.'
                inc eax
                mov byte ptr [eax+ecx],'*'
                inc eax
                mov byte ptr [eax+ecx],00h
                Invoke WildCardCopy,pszSourcePath,pszTargetPath,dSubFolders,dOverWrite,TRUE
            .else
                ;*************** Process multiple wildcard separated by semicolon ";"
                Invoke lstrlen,pszSourcePath
                dec eax
                mov ecx,pszSourcePath
                @@:
                cmp byte ptr [ecx+eax],05Ch ;'\'
                je @F
                    dec eax
                    jmp @B
                @@:
                add eax,1
                add eax,1
                push eax
                Invoke lstrcpyn,addr szBaseWildcard,pszSourcePath,eax
                mov eax,pszSourcePath
                pop ecx
                sub ecx,1
                add eax,ecx
                Invoke lstrcpy,addr szTrailWildcard,eax
                ;*************** Break WildCardCopy for every semicolon ";"
                Invoke lstrcpy,addr szSingleWildcard,addr szBaseWildcard
                Invoke lstrlen,addr szTrailWildcard
                lea ecx,szTrailWildcard
                dec eax
                NextWildCard:
                    cmp byte ptr [ecx+eax],';'
                    jne @F
                        push eax
                        push ecx
                        ;
                        mov byte ptr [ecx+eax],0
                        add eax,1
                        add eax,ecx
                        Invoke lstrcat,addr szSingleWildcard,eax
                        Invoke WildCardCopy,addr szSingleWildcard,pszTargetPath,dSubFolders,dOverWrite,TRUE
                        Invoke lstrcpy,addr szSingleWildcard,addr szBaseWildcard
                        ;
                        pop ecx
                        pop eax
                    @@:
                    dec eax
                    jnz NextWildCard
                Invoke lstrcat,addr szSingleWildcard,addr szTrailWildcard
                Invoke WildCardCopy,addr szSingleWildcard,pszTargetPath,dSubFolders,dOverWrite,TRUE
            .endif
        .endif
          ;mov CurrentFile,0
        ExitCopyFiles:
        ;
    ret
CopyFiles                       endp
;########################################################################
;                 *** CopyFiles Procedure ***
;#########################################################################
align 4
WildCardCopy                    proc pSourceFolder:DWORD,pDestFolder:DWORD,dSubFolders:DWORD,dOverWrite:DWORD,dIsParent:DWORD
                                ;=====================
                                ; Put LOCALs on stack
                                ;=====================
                                LOCAL Win32FindData:WIN32_FIND_DATA
                                LOCAL hFindFile:DWORD
                                LOCAL szPattern[MAX_PATH]:BYTE
                                LOCAL szDirBase[MAX_PATH]:BYTE
                                LOCAL szDirDest[MAX_PATH]:BYTE
                                LOCAL szDirRecursive[MAX_PATH]:BYTE

        Invoke DirExist,pDestFolder                                      ;Creates directory of dest. file if NO exists.
        Invoke FindFirstFile,pSourceFolder,addr Win32FindData
        mov hFindFile, eax                                              ;Save Handle in esi register
        inc eax                                                         ;INVALID_HANDLE_VALUE = -1
        je InvalidHandle
        CheckForFiles:
            cmp eax,ERROR_NO_MORE_FILES
            je LeaveLoop
            .if byte ptr Win32FindData.cFileName == '.'
                jmp Continue
            .endif
            Invoke DirFromPath,pSourceFolder,addr szDirBase,addr szPattern
            Invoke lstrcat, addr szDirBase,addr Win32FindData.cFileName
            .if (Win32FindData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
                .if dSubFolders == TRUE
                    Invoke lstrcpy,addr szDirDest,pDestFolder
                    Invoke lstrcat,addr szDirDest,addr Win32FindData.cFileName
                    Invoke lstrlen,addr szDirDest
                    lea ecx,szDirDest
                    add ecx,eax
                    mov byte ptr [ecx],'\'
                    inc ecx
                    mov byte ptr [ecx],00h
                    .if dIsParent==TRUE
                        Invoke lstrcpy,addr szDirRecursive,addr szDirDest
                        Invoke lstrlen,addr szDirRecursive
                        push eax
                        Invoke lstrlen,addr szDirBase
                        pop ecx
                        .if ecx>eax ;If szDirRecursive is greater szDirBase
                            lea ecx,szDirRecursive
                            add ecx,eax
                            mov byte ptr [ecx],00h ;Truncate szDirRecursive Path to the same Len or szDirBase
                        .endif
                        Invoke CompareString, LOCALE_USER_DEFAULT, NORM_IGNORECASE,addr szDirBase, -1,addr szDirRecursive, -1
                        .if eax==2 ;2   The string pointed to by lpString1 is equal in lexical value to the string pointed to by lpString2.
                            ;Invoke MessageBox,NULL,addr szDirBase,addr szDirRecursive,NULL
                            jmp Continue ;This will generate a Recursive Copy, so ommit this directory.
                        .endif
                    .endif
                    Invoke lstrcat,addr szDirBase,addr szPattern
                    Invoke WildCardCopy,addr szDirBase,addr szDirDest,dSubFolders,dOverWrite,FALSE
                    Invoke SetCurrentDirectory,addr szDirBase                     ;Restores Previus CurrentDirectory
                .endif
                jmp Continue
            .endif
            Invoke SingleCopy,addr szDirBase,pDestFolder,dOverWrite
            Continue:
                invoke FindNextFile,hFindFile,addr Win32FindData
                invoke GetLastError  ;cmp eax & je LeaveLoop,ERROR_NO_MORE_FILES ;Validate on CheckForFile
                jmp CheckForFiles
        LeaveLoop:
        invoke FindClose,hFindFile
        InvalidHandle:
    ret

WildCardCopy                    endp
;#########################################################################
;                 *** CopyFiles Procedure ***
;#########################################################################
align 4
SingleCopy                      proc SourceFile:DWORD,TargeFile:DWORD,dOverWrite:DWORD  ;SourceFile Must have a "full path"
                                ;
                                LOCAL szFullTargetFile[MAX_PATH]:BYTE
                                LOCAL lpTargetFile:DWORD
        ;
        ;Invoke MessageBox,NULL,TargeFile,SourceFile,MB_OK
        mov eax,TargeFile
        mov lpTargetFile,eax
        Invoke lstrlen,TargeFile
        dec eax                                           ;Decrement the Null char
        mov ecx,eax                                       ;Copy Len to ecx
        mov eax,TargeFile                                 ;Copy address of Target File to eax
        add ecx,eax                                       ;Set pointer on ecx to the last char
        mov al,[ecx]
        cmp al,'\'
        jnz CopyFullNamePath
            Invoke lstrcpy,addr szFullTargetFile,TargeFile
            Invoke lstrlen,SourceFile                     ;Len of SourceFile
            dec eax                                       ;Decrement the Null char
            mov ecx,eax                                   ;Copy Len to ecx
            mov eax,SourceFile                            ;Copy address of Target File to eax
            add ecx,eax                                   ;Set pointer on ecx to the last char
            @@:
            dec ecx
            mov al,[ecx]
            cmp al,'\'
            jne @B
            inc ecx                                       ;In ecx got the pointer of realfile name vgr: file.exe
            Invoke lstrcat,addr szFullTargetFile,ecx      ;lstrcpy uses ecx,ebx and eax registers
            lea eax,szFullTargetFile
            mov lpTargetFile,eax
        CopyFullNamePath:
        Invoke DirExist,lpTargetFile
        .if dOverWrite
            xor ecx,ecx
        .else
            mov ecx,COPY_FILE_FAIL_IF_EXISTS
        .endif
        ;
        Invoke CopyFileEx,SourceFile,lpTargetFile,NULL,NULL,NULL,ecx
        ;
    ret
SingleCopy                      endp
;########################################################################
;                 *** DirExist Procedure ***
;#########################################################################
align 4
DirExist                        proc  pDestFolder:DWORD
                                ; Create Directory If Does NOT exist and Set it as Current Directory
                                LOCAL CurDir[MAX_PATH]:BYTE
                                LOCAL PathDir[MAX_PATH]:BYTE
        ;
        ;Invoke MessageBox,NULL,DestFolder,DestFolder,MB_OK
        Invoke DirFromPath,pDestFolder,addr PathDir,NULL                ;Extract last Root Directory from DestFolder
        Invoke GetCurrentDirectory,MAX_PATH,addr CurDir                 ;Saves a temporal CurrentDirectory
        Invoke SetCurrentDirectory,addr PathDir
        test eax,eax
        jne @F
            Invoke CreateDirectory,addr PathDir,NULL                    ;If Directory does not exist then create it
        @@:
        Invoke SetCurrentDirectory,addr CurDir                          ;Restores a temporal CurrentDirectory
    ret
    ;
DirExist                        endp
;########################################################################
;                 *** DirExist Procedure ***
;#########################################################################
align 4
FileExist                       proc  pszFilePath:DWORD
                                ;
                                LOCAL wfd:WIN32_FIND_DATA
        ;                       ;
        Invoke FindFirstFile,pszFilePath,addr wfd
        .if eax != INVALID_HANDLE_VALUE ; File exists
            xor eax,eax
            add eax,1
        .else
            xor eax,eax
        .endif
    ret
    ;
FileExist                       endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;BuildTreeDir
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
BuildTreeDir                    proc  pDestFolder:DWORD
                                ;Test if Dir Exits (And Create if Dont)
        ;
        push edi
        push esi
        mov edi,pDestFolder
        cmp byte ptr [edi],5Ch  ;Skip \ if it's the first character
        jz Foo
        cmp byte ptr [edi+1],':' ;Skip X:\
        jnz Again
        add edi,3 ;

    Again:
        movzx ESI,byte ptr [edi]

        .if esi==0 || esi==5Ch

            mov byte ptr [edi],0    ;Nullterminate where \ is found
            invoke CreateDirectory,pDestFolder,NULL
            invoke GetFileAttributes,pDestFolder
            mov byte ptr [edi],5Ch

            inc eax         ;Does the current directory exist?
            or eax,eax
            ;jnz @F
            ;jmp Exit       ;Error: Return 0 and exit
            ;@@:
            jz Exit

            or ESI,ESI      ;End of path? :-P
            jnz @F
            mov byte ptr [edi],0    ;Put the 0h back at the end
            mov eax,1
            jmp Exit        ;Success: Return 1 and exit
            @@:

        .endif

    Foo:
        inc edi
        jmp Again


    Exit:
        push esi
        push edi
        ;
        xor eax,eax
    ret
    ;
BuildTreeDir                    endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;Extracs (C:\Temp\) From C:\Temp\file.txt
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
DirFromPath                     proc SourcePath:DWORD, TargetBuffer:DWORD,  LeaderBuffer:DWORD   ;Origen,Target,Leader Extracs (C:\Temp\) From C:\Temp\file.txt
                        ;=====================
                        ; Put LOCALs on stack
                        ;=====================
                        LOCAL szOnlyDirectoryPath[MAX_PATH]:BYTE
                        LOCAL lpSourcePath:DWORD
                        LOCAL lpSlash:DWORD

      mov eax,SourcePath
      mov lpSourcePath,eax
      Invoke lstrlen,SourcePath
      dec eax                                           ;Decrement the Null char
      mov ecx,eax                                       ;Copy Len to ecx
      mov eax,SourcePath                                ;Copy address of SourcePath to eax
      add ecx,eax                                       ;Set pointer on ecx to the last char
      mov al,[ecx]
      cmp al,'\'
      jz IsDirectory
          ;Pendiente: Aqui debe Verificar que la Ruta NO sea un directorio sin el Slash "\"
          @@:
          dec ecx
          mov al,[ecx]
          cmp al,'\'
          jne @B
          mov lpSlash,ecx
          add ecx,2                                     ;add space to the previous "\" and to null character
          mov eax,SourcePath
          sub ecx,eax
          Invoke lstrcpyn,addr szOnlyDirectoryPath,SourcePath,ecx   ;copy till the "\" position
          lea eax,szOnlyDirectoryPath
          mov lpSourcePath,eax
          Invoke lstrcpy,LeaderBuffer,lpSlash ;,ecx   ;copy till the "\" position
      IsDirectory:
      Invoke lstrcpy,TargetBuffer,lpSourcePath
      ret
DirFromPath                     endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;AscciiToDword  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
AscciiToDword                   proc pszCstring:DWORD
                                ;
        ; ----------------------------------------
        ; Convert decimal string into dword value
        ; return value in eax
        ; ----------------------------------------

        push esi
        push edi

        xor eax, eax
        mov esi, [pszCstring]
        xor ecx, ecx
        xor edx, edx
        mov al, [esi]
        inc esi
        cmp al, 2D
        jne proceed
            mov al, byte ptr [esi]
            not edx
            inc esi
        jmp proceed

        @@:
            sub al, 30h
            lea ecx, dword ptr [ecx+4*ecx]
            lea ecx, dword ptr [eax+2*ecx]
            mov al, byte ptr [esi]
            inc esi
        proceed:
        or al, al
        jne @B
        lea eax, dword ptr [edx+ecx]
        xor eax, edx

        pop edi
        pop esi
        ;
    ret
AscciiToDword                   endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;DwordToAsccii  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
DwordToAsccii                   proc dwValue:DWORD,pBuffer:DWORD
                                ; -------------------------------------------------------------
                                ; convert DWORD to ascii string
                                ; dwValue is value to be converted
                                ; lpBuffer is the address of the receiving buffer
                                ; EXAMPLE:
                                ; invoke dwtoa,edx,ADDR buffer
                                ;
                                ; Uses: eax, ecx, edx.
                                ; -------------------------------------------------------------
        ;
        push ebx
        push esi
        push edi

        mov eax, dwValue
        mov edi, [pBuffer]

        test eax,eax
        jnz sign

      zero:
        mov word ptr [edi],30h
        jmp dtaexit

      sign:
        jns pos
        mov byte ptr [edi],'-'
        neg eax
        add edi, 1

      pos:
        mov ecx, 3435973837
        mov esi, edi

        .while (eax > 0)
          mov ebx,eax
          mul ecx
          shr edx, 3
          mov eax,edx
          lea edx,[edx*4+edx]
          add edx,edx
          sub ebx,edx
          add bl,'0'
          mov [edi],bl
          add edi, 1
        .endw

        mov byte ptr [edi], 0       ; terminate the string

        ; We now have all the digits, but in reverse order.

        .while (esi < edi)
          sub edi, 1
          mov al, [esi]
          mov ah, [edi]
          mov [edi], al
          mov [esi], ah
          add esi, 1
        .endw

      dtaexit:

        pop edi
        pop esi
        pop ebx
        ;
    ret
    ;
DwordToAsccii                   endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;MemCopy
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
MemCopy                         proc dwSource:DWORD,dwTarget:DWORD,dwSize:DWORD
;MemCopy                         proc uses esi edi dwSource:DWORD,dwTarget:DWORD,dwSize:DWORD
;Original MemCopy                         proc public uses esi edi dwSource:PTR BYTE,dwTarget:PTR BYTE,dwSize:DWORD
                                ;
        ;
        Invoke RtlMoveMemory,dwTarget,dwSource,dwSize
        ;
;        ; ---------------------------------------------------------
;        ; Copy ln bytes of memory from Source buffer to Dest buffer
;        ;      ~~                      ~~~~~~           ~~~~
;        ; USAGE:
;        ; invoke MemCopy,ADDR Source,ADDR Dest,4096
;        ;
;        ; NOTE: Dest buffer must be at least as large as the source
;        ;       buffer otherwise a page fault will be generated.
;        ; ---------------------------------------------------------
;        ;
;        cld
;        mov esi,[dwSource]
;        mov edi,[dwTarget]
;        mov ecx,[dwSize]
;        ;
;        shr ecx,2
;        rep movsd
;        ;
;        mov ecx,[dwSize]
;        and ecx, 3
;        rep movsb
;        xor eax,eax
;        ;
    ret
    ;
MemCopy                         endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;FillBuffer  PROCEDURE
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
FillBuffer                      proc lpBuffer:DWORD,lenBuffer:DWORD,TheChar:BYTE
        ;
        push edi
        ;
        mov edi, lpBuffer   ; address of buffer
        mov ecx, lenBuffer  ; buffer length
        mov  al, TheChar    ; load al with character
        rep stosb           ; write character to buffer until ecx = 0
        ;
        pop edi
        ;
    ret
FillBuffer                      endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;StrLen
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
StrLen                          proc item:DWORD

  ; -------------------------------------------------------------
  ; This procedure has been adapted from an algorithm written by
  ; Agner Fog. It has the unusual characteristic of reading up to
  ; three bytes past the end of the buffer as it uses DWORD size
  ; reads. It is measurably faster than a classic byte scanner on
  ; large linear reads and has its place where linear read speeds
  ; are important.
  ; -------------------------------------------------------------

    push    ebx
    mov     eax,item               ; get pointer to string
    lea     edx,[eax+3]            ; pointer+3 used in the end
  @@:
    mov     ebx,[eax]              ; read first 4 bytes
    add     eax,4                  ; increment pointer
    lea     ecx,[ebx-01010101h]    ; subtract 1 from each byte
    not     ebx                    ; invert all bytes
    and     ecx,ebx                ; and these two
    and     ecx,80808080h
    jz      @B                     ; no zero bytes, continue loop
    test    ecx,00008080h          ; test first two bytes
    jnz     @F
    shr     ecx,16                 ; not in the first 2 bytes
    add     eax,2
  @@:
    shl     cl,1                   ; use carry flag to avoid branch
    sbb     eax,edx                ; compute length
    pop     ebx

    ret

StrLen                          endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;InString
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
InString                        proc startpos:DWORD,lpSource:DWORD,lpPattern:DWORD

  ; ------------------------------------------------------------------
  ; InString searches for a substring in a larger string and if it is
  ; found, it returns its position in eax.
  ;
  ; It uses a one (1) based character index (1st character is 1,
  ; 2nd is 2 etc...) for both the "StartPos" parameter and the returned
  ; character position.
  ;
  ; Return Values.
  ; If the function succeeds, it returns the 1 based index of the start
  ; of the substring.
  ;  0 = no match found
  ; -1 = substring same length or longer than main string
  ; -2 = "StartPos" parameter out of range (less than 1 or longer than
  ; main string)
  ; ------------------------------------------------------------------

    LOCAL sLen:DWORD
    LOCAL pLen:DWORD

    push ebx
    push esi
    push edi

    invoke StrLen,lpSource
    mov sLen, eax           ; source length
    invoke StrLen,lpPattern
    mov pLen, eax           ; pattern length

    cmp startpos, 1
    jge @F
    mov eax, -2
    jmp isOut               ; exit if startpos not 1 or greater
  @@:

    dec startpos            ; correct from 1 to 0 based index

    cmp  eax, sLen
    jl @F
    mov eax, -1
    jmp isOut               ; exit if pattern longer than source
  @@:

    sub sLen, eax           ; don't read past string end
    inc sLen

    mov ecx, sLen
    cmp ecx, startpos
    jg @F
    mov eax, -2
    jmp isOut               ; exit if startpos is past end
  @@:

  ; ----------------
  ; setup loop code
  ; ----------------
    mov esi, lpSource
    mov edi, lpPattern
    mov al, [edi]           ; get 1st char in pattern

    add esi, ecx            ; add source length
    neg ecx                 ; invert sign
    add ecx, startpos       ; add starting offset

    jmp Scan_Loop

    align 16

  ; @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

  Pre_Scan:
    inc ecx                 ; start on next byte

  Scan_Loop:
    cmp al, [esi+ecx]       ; scan for 1st byte of pattern
    je Pre_Match            ; test if it matches
    inc ecx
    js Scan_Loop            ; exit on sign inversion

    jmp No_Match

  Pre_Match:
    lea ebx, [esi+ecx]      ; put current scan address in EBX
    mov edx, pLen           ; put pattern length into EDX

  Test_Match:
    mov ah, [ebx+edx-1]     ; load last byte of pattern length in main string
    cmp ah, [edi+edx-1]     ; compare it with last byte in pattern
    jne Pre_Scan            ; jump back on mismatch
    dec edx
    jnz Test_Match          ; 0 = match, fall through on match

  ; @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

  Match:
    add ecx, sLen
    mov eax, ecx
    inc eax
    jmp isOut

  No_Match:
    xor eax, eax

  isOut:
    pop edi
    pop esi
    pop ebx

    ret

InString                        endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; #########################################################################
                    ;InString
; #########################################################################
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
align 4
GetPathOnly                     proc src:DWORD, dst:DWORD

        push esi
        push edi

        xor ecx, ecx    ; zero counter
        mov esi, src
        mov edi, dst
        xor edx, edx    ; zero backslash location

        @@:
        mov al, [esi]   ; read byte from address in esi
        inc esi
        inc ecx         ; increment counter
        cmp al, 0       ; test for zero
            je gfpOut       ; exit loop on zero
            cmp al, "\"     ; test for "\"
            jne nxt1        ; jump over if not
            mov edx, ecx    ; store counter in ecx = last "\" offset in ecx
        nxt1:
            mov [edi], al   ; write byte to address in edi
            inc edi
            jmp @B

        gfpOut:
            add edx, dst    ; add destination address to offset of last "\"
            mov [edx], al   ; write terminator to destination

        pop edi
        pop esi
        ;
    ret
    ;
GetPathOnly                     endp
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
; ллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллллл
end start