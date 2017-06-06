Attribute VB_Name = "funcs"
Option Explicit

Private Declare Function DnsQuery Lib "dnsapi" Alias "DnsQuery_A" (ByVal strname As String, ByVal wType As Integer, ByVal fOptions As Long, ByVal pServers As Long, ppQueryResultsSet As Long, ByVal pReserved As Long) As Long
Private Declare Function DnsRecordListFree Lib "dnsapi" (ByVal pDnsRecord As Long, ByVal FreeType As Long) As Long
Private Declare Function lstrlen Lib "kernel32" (ByVal straddress As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal pIP As Long) As Long
Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal sAddr As String) As Long
Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long


Private Const DnsFreeRecordList         As Long = 1
Private Const DNS_TYPE_A                As Long = &H1
Private Const DNS_QUERY_BYPASS_CACHE    As Long = &H8

Private Type VBDnsRecord
    pNext           As Long
    pName           As Long
    wType           As Integer
    wDataLength     As Integer
    flags           As Long
    dwTel           As Long
    dwReserved      As Long
    prt             As Long
    others(35)      As Byte
End Type

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Sub MakeNormal(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub





Public Function Resolve(sAddr As String, Optional sDnsServers As String) As String
    Dim pRecord     As Long
    Dim pNext       As Long
    Dim uRecord     As VBDnsRecord
    Dim lPtr        As Long
    Dim vSplit      As Variant
    Dim laServers() As Long
    Dim pServers    As Long
    Dim sName       As String

    If LenB(sDnsServers) <> 0 Then
        vSplit = Split(sDnsServers)
        ReDim laServers(0 To UBound(vSplit) + 1)
        laServers(0) = UBound(laServers)
        For lPtr = 0 To UBound(vSplit)
            laServers(lPtr + 1) = inet_addr(vSplit(lPtr))
        Next
        pServers = VarPtr(laServers(0))
    End If
    If DnsQuery(sAddr, DNS_TYPE_A, DNS_QUERY_BYPASS_CACHE, pServers, pRecord, 0) = 0 Then
        pNext = pRecord
        Do While pNext <> 0
            Call CopyMemory(uRecord, pNext, Len(uRecord))
            If uRecord.wType = DNS_TYPE_A Then
                lPtr = inet_ntoa(uRecord.prt)
                sName = String(lstrlen(lPtr), 0)
                Call CopyMemory(ByVal sName, lPtr, Len(sName))
                If LenB(Resolve) <> 0 Then
                    Resolve = Resolve & " "
                End If
                Resolve = Resolve & sName
            End If
            pNext = uRecord.pNext
        Loop
        Call DnsRecordListFree(pRecord, DnsFreeRecordList)
    End If
End Function



Public Function openInBrowser(url As String)
    'CreateObject("Wscript.Shell").Run url
     Dim r As Long
     r = ShellExecute(0, "open", url, 0, 0, 1)
     'MsgBox (r)
End Function




