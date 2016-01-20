
Option Explicit

Private Declare PtrSafe Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long

Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const STATUS_PENDING = &H103&
Private Const PROCESS_QUERY_INFORMATION = &H400



Public Function ShellandWait(ExeFullPath As String, _
Optional TimeOutValue As Long = 0) As Boolean
    
    Dim lInst As Long
    Dim lStart As Long
    Dim lTimeToQuit As Long
    Dim sExeName As String
    Dim lProcessId As Long
    Dim lExitCode As Long
    Dim bPastMidnight As Boolean
    
    
    On Error GoTo ErrorHandler

    lStart = CLng(Timer)
    sExeName = ExeFullPath

    'Deal with timeout being reset at Midnight
    If TimeOutValue > 0 Then
        If lStart + TimeOutValue < 86400 Then
            lTimeToQuit = lStart + TimeOutValue
        Else
            lTimeToQuit = (lStart - 86400) + TimeOutValue
            bPastMidnight = True
        End If
    End If

    lInst = Shell(sExeName, vbMinimizedNoFocus)
    
    lProcessId = OpenProcess(PROCESS_QUERY_INFORMATION, False, lInst)

    Do
        Call GetExitCodeProcess(lProcessId, lExitCode)
        DoEvents
        If TimeOutValue And Timer > lTimeToQuit Then
            If bPastMidnight Then
                 If Timer < lStart Then Exit Do
            Else
                 Exit Do
            End If
    End If
    Loop While lExitCode = STATUS_PENDING
    
    ShellandWait = True
   
ErrorHandler:
ShellandWait = False
Exit Function
End Function



Private Sub Cancel_Button_Click()
    Unload Me
End Sub

Private Sub Ok_Button_Click()
    Dim fso As Object
    Dim texout
    Dim sOld As Shape
    Dim s As String
    Dim path As String
    Dim curpath As String
    Dim d11 As Double, d12 As Double, d21 As Double, d22 As Double
    Dim tx As Double, ty As Double
    Dim fsT As Object
    Dim objStreamUTF8NoBOM: Set objStreamUTF8NoBOM = CreateObject("ADODB.Stream") ' Just to remove BOM from utf-8 encoding
    Set fsT = CreateObject("ADODB.Stream")
    Set sOld = ActiveShape
    If Not (sOld Is Nothing) Then
        sOld.GetMatrix d11, d12, d21, d22, tx, ty
    End If
    path = Environ$("TEMP")
    curpath = CurDir
    s = TextBox1.Text
    path = path & "\"

fsT.Type = 2 'Specify stream type - we want To save text/string data.
fsT.CharSet = "utf-8" 'Specify charset For the source text data. Error becouse of BOM
'fsT.LineSeparator = -1
fsT.Open 'Open the stream And write binary data To the object

fsT.WriteText "% UFT-8 file for use of Portuguese accent in Corel ", 1
fsT.WriteText "\documentclass[10.5pt,a5paper,english,brazil]{abntex2}", 1
fsT.WriteText "\usepackage[T1]{fontenc}", 1
fsT.WriteText "\usepackage[utf8]{inputenc}", 1
fsT.WriteText "\usepackage{amsmath}", 1
fsT.WriteText "\usepackage{amsfonts}", 1
fsT.WriteText "\usepackage{lmodern}", 1
fsT.WriteText "\usepackage{siunitx}", 1
fsT.WriteText "\usepackage{tikz,pgfplots}", 1
fsT.WriteText "\usepackage{anyfontsize}", 1
fsT.WriteText "\usepackage[americancurrents, americanvoltages, americanresistors, cuteinductors,americanports, siunitx, noarrowmos, smartlabels]{circuitikz}", 1
fsT.WriteText "\usepackage{booktabs}", 1
fsT.WriteText "\renewcommand{\rmdefault}{cmr} % Selects a roman (i.e., serifed) font family", 1
fsT.WriteText "\renewcommand{\familydefault}{cmr} % Fonte padrão utilizada no texto  ", 1
fsT.WriteText "\renewcommand{\normalsize}{\fontsize{10.5pt}{11pt}}", 1
fsT.WriteText "\sisetup{detect-all}", 1
fsT.WriteText "\sisetup{scientific-notation = fixed, fixed-exponent = 0, round-mode = places,round-precision = 2,output-decimal-marker = {,} }", 1
fsT.WriteText "% Begin of document ", 1
fsT.WriteText "\begin{document}", 1
fsT.WriteText "\thispagestyle{empty}", 1
fsT.WriteText "\fontsize{9.5pt}{10.5pt}\selectfont{", 1
fsT.WriteText s, 1
fsT.WriteText "}\end{document}"

  fsT.Position = 3
  objStreamUTF8NoBOM.Type = 1
  objStreamUTF8NoBOM.Open
  fsT.CopyTo objStreamUTF8NoBOM
  objStreamUTF8NoBOM.SaveToFile (path + "teximport.tex"), 2
  
  
 
 
    ChDrive path
    ChDir path
    ShellandWait ("latex.exe """ + "teximport.tex""")
    ShellandWait ("dvips.exe """ + "teximport.dvi""")
    ChDrive curpath
    ChDir curpath
    Dim impflt As ImportFilter
    Dim impopt As StructImportOptions
    Set impopt = New StructImportOptions
    impopt.MaintainLayers = True
    Set impflt = ActiveLayer.ImportEx(path + "teximport.ps", cdrPSInterpreted, impopt)
    impflt.Finish
    Dim s1 As Shape
    Set s1 = ActiveShape
    Rem s1.SetPosition sOld.PositionX, sOld.PositionY
    If Not (sOld Is Nothing) Then
        s1.SetMatrix d11, d12, d21, d22, tx, ty
        sOld.Delete
    End If
    s1.Name = s
    s1.ObjectData("Comments") = s
    Unload Me
End Sub

