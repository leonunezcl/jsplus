VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInsertar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ssi_break_html() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#break -->"
    
    ssi_break_html = ret.ToString
    
    Set ret = Nothing
    
End Function

Public Function ssi_echo() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#echo var=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_echo = ret.ToString
    
    Set ret = Nothing
    
End Function


Public Function ssi_error_msg() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#config errmsg=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_error_msg = ret.ToString
    
    Set ret = Nothing
    
End Function

Public Function ssi_exec_cgi() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#exec cgi=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_exec_cgi = ret.ToString
    
    Set ret = Nothing
    
End Function


Public Function ssi_exec_cmd() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#exec cmd=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_exec_cmd = ret.ToString
    
    Set ret = Nothing
    
End Function


Public Function ssi_file_last_mod() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#flastmod file=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_file_last_mod = ret.ToString
    
    Set ret = Nothing
    
End Function


Public Function ssi_file_last_mod_virtual() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#flastmod virtual=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_file_last_mod_virtual = ret.ToString
    
    Set ret = Nothing
    
End Function


Public Function ssi_file_size() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#fsize file=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_file_size = ret.ToString
    
    Set ret = Nothing
    
End Function


Public Function ssi_file_size_virtual() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#fsize virtual=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_file_size_virtual = ret.ToString
    
    Set ret = Nothing
    
End Function


Public Function ssi_goto() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#goto=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_goto = ret.ToString
    
    Set ret = Nothing
    
End Function


Public Function ssi_goto_label() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#label=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_goto_label = ret.ToString
    
    Set ret = Nothing
    
End Function


Public Function ssi_if() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#if expr=" & Chr$(34) & Chr$(34) & " -->" & vbNewLine
    ret.Append "" & vbNewLine
    ret.Append "<!--#endif -->"
    
    ssi_if = ret.ToString
    
    Set ret = Nothing
    
End Function

Public Function ssi_if_else() As String
            
    Dim ret As New cStringBuilder
    
    ret.Append "<!--#if expr=" & Chr$(34) & Chr$(34) & " -->" & vbNewLine
    ret.Append "" & vbNewLine
    ret.Append "<!--#else -->" & vbNewLine
    ret.Append "" & vbNewLine
    ret.Append "<!--#endif -->"
    
    ssi_if_else = ret.ToString
    
    Set ret = Nothing
    
End Function


Public Function ssi_include_file() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#include file=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_include_file = ret.ToString
    
    Set ret = Nothing
        
End Function


Public Function ssi_include_virtual() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#include virtual=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_include_virtual = ret.ToString
    
    Set ret = Nothing
    
End Function


Public Function ssi_set_variable() As String

    Dim ret As New cStringBuilder
    
    ret.Append "<!--#set var=" & Chr$(34) & Chr$(34) & " value=" & Chr$(34) & Chr$(34) & " -->"
    
    ssi_set_variable = ret.ToString
    
    Set ret = Nothing
    
End Function


