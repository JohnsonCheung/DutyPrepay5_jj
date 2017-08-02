Attribute VB_Name = "e"
Option Compare Text
Option Explicit

'============================
'============================
'============================
'============================
'============================
' Enum
'============================
'============================
'============================
'============================
'============================
Public Enum eTypImpCurVal
    eTblF = 1
    
End Enum

Public Enum eTypFct
    eSub = 1
    eFct = 2
    eGet = 3
    eLet = 4
    eSet = 5
End Enum
Public Enum eHostSts
    e1Rec = 1
    e0Rec = 2
    e2Rec = 3
    eHostCpyToFrm = 4
    eUnExpectedErr = 5
End Enum
Public Enum eEdtMd
    Add = 1
    Edt = 2
    Dlt = 4
    Sel = 8
    Mtc = 7
End Enum
Public Enum eLang
    eEN = 1 ' English
    eTC = 2 ' Traditional Chinese
    eSC = 3 ' Simplified Chinese
End Enum
Public Enum eTypMsg
    ePrmErr = 1
    eCritical = 2
    eTrc = 3
    eWarning = 4
    eSeePrvMsg = 5
    eException = 6
    eUsrInfo = 7
    eRunTimErr = 8
    eImpossibleReachHere = 9
    eQuit = 10
End Enum
Public Enum eTimStampOpt
    eNoStamp = 0
    eYr = 1
    eMth = 2
    eWk = 3
    eDte = 4
    eMin = 5
End Enum
'[x] [%x:x] [x:x:x] [>x] [>=x] [<x] [<=x] [*x] [x*] [*x*] [!%x:x] [!x:x:x] [!*x] [!x*] [!*x*] [!x]
Public Enum eOpTyp
    eEq = 1
    eRge = 2
    eLst = 3
    eGt = 4
    eGe = 5
    eLt = 6
    Ele = 7
    eLik = 8
    eNRge = 9
    eNLst = 10
    eNLik = 11
    eNe = 12
End Enum
Public Enum eTypSim
    eTypSim_Num = 1
    eTypSim_Str = 2
    eTypSim_Bool = 3
    eTypSim_Dte = 4
    eTypSim_Oth = 5
End Enum
'Public Enum eBrkStrOpt
'    eNoTrim = 1      'Do not trim S1 & S2
'    eIs1or2 = 2      'Is 1 or 2 elemens? error in no item.
'    eIs2 = 4         'Is 2 element?      Error in no or one item.
''   OneForS1 = 0    'If only 1 element is given, assign this element to S1 and set S2 to blank
'    e1ForS2 = 8    'If only 1 element is given, assign this element to S2 and set S1 to blank
'    e1ForBoth = 16 'If only 1 element is given, assign this element to both S1 & S2
'End Enum
