# Tulip Cell
# https://tulipcell.org/
# Copyright (c) 2010-2016 Tulip Charts LLC
# Lewis Van Winkle (LV@tulipcharts.org)
#
# This file is part of Tulip Cell.
#
# Tulip Cell is free software: you can redistribute it and/or modify it
# under the terms of the GNU Lesser General Public License as published by the
# Free Software Foundation, either version 3 of the License, or (at your
# option) any later version.
#
# Tulip Cell is distributed in the hope that it will be useful, but
# WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
# FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser General Public License
# for more details.
#
# You should have received a copy of the GNU Lesser General Public License
# along with Tulip Cell.  If not, see <http://www.gnu.org/licenses/>.


#This program creates the VBA interface to excel for tulipcell.dll


set version 0.8.1
set version_number 801

#Set this to the path for Tulip Indicators sample.exe program
set sample_path {../tulipindicators/sample.exe}

set ti_version [exec $sample_path --version]



set out [open vba.txt w]
fconfigure $out -translation {auto lf}

puts $out {'/*
' * Tulip Cell
' * https://tulipcell.org/
' * Copyright (c) 2010-2016 Tulip Charts LLC
' * Lewis Van Winkle (LV@tulipcharts.org)
' *
' * This file is part of Tulip Cell.
' *
' * Tulip Cell is free software: you can redistribute it and/or modify it
' * under the terms of the GNU Lesser General Public License as published by the
' * Free Software Foundation, either version 3 of the License, or (at your
' * option) any later version.
' *
' * Tulip Cell is distributed in the hope that it will be useful, but
' * WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
' * FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser General Public License
' * for more details.
' *
' * You should have received a copy of the GNU Lesser General Public License
' * along with Tulip Cell.  If not, see <http://www.gnu.org/licenses/>.
' *
' */
}

puts $out "' * Tulip Cell version: $version"
puts $out "' * $ti_version"

puts $out {


'ADD THIS CODE TO THE WORKBOOK
Private WithEvents App As Application

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    TI_RegisterHelp
    TI_CheckForUpdate
End Sub

Private Sub Workbook_Open()
    Set App = Application
End Sub
'END OF WORKBOOK CODE



Option Explicit

#If win64 Then
Private Declare PtrSafe Function TI_GetIndicator Lib "tulipcell64.dll" Alias "GetIndicator" (ByVal name As String) As Long
Private Declare PtrSafe Function TI_Call Lib "tulipcell64.dll" Alias "Call" (ByVal index As Long, ByVal size As Long, ByRef inputs_in As Double, ByRef options As Double, ByRef outputs As Double) As Long
Private Declare PtrSafe Function TI_GetStart Lib "tulipcell64.dll" Alias "GetStart" (ByVal index As Long, ByRef options As Double) As Long
Private Declare PtrSafe Function TI_GetInputCount Lib "tulipcell64.dll" Alias "GetInputCount" (ByVal index As Long) As Long
Private Declare PtrSafe Function TI_GetOptionCount Lib "tulipcell64.dll" Alias "GetOptionCount" (ByVal index As Long) As Long
Private Declare PtrSafe Function TI_GetOutputCount Lib "tulipcell64.dll" Alias "GetOutputCount" (ByVal index As Long) As Long
#Else
Private Declare Function TI_GetIndicator Lib "tulipcell32.dll" Alias "GetIndicator" (ByVal name As String) As Long
Private Declare Function TI_Call Lib "tulipcell32.dll" Alias "Call" (ByVal index As Long, ByVal size As Long, ByRef inputs_in As Double, ByRef options As Double, ByRef outputs As Double) As Long
Private Declare Function TI_GetStart Lib "tulipcell32.dll" Alias "GetStart" (ByVal index As Long, ByRef options As Double) As Long
Private Declare Function TI_GetInputCount Lib "tulipcell32.dll" Alias "GetInputCount" (ByVal index As Long) As Long
Private Declare Function TI_GetOptionCount Lib "tulipcell32.dll" Alias "GetOptionCount" (ByVal index As Long) As Long
Private Declare Function TI_GetOutputCount Lib "tulipcell32.dll" Alias "GetOutputCount" (ByVal index As Long) As Long
#End If


Dim TI_HasUpdate as Integer

Public Sub TI_CheckForUpdate()
    On Error GoTo errHandler

    If (TI_HasUpdate <> 0) then
        goto done
    End If

    Dim ie As Object
    Set ie = CreateObject("internetexplorer.application")
    ie.Visible = False

    Dim version as long}

puts $out "    version = $version_number"

puts $out {
    ie.navigate "https://tulipcell.org/update?version=" & version & "&extra=" & Application.Version

    Do While ie.readystate <> 4: DoEvents: Loop

    Dim html As String
    html = ie.Document.DocumentElement.innerHTML()

    If html Like "*update ready*" Then
        TI_HasUpdate = 1
    else
        TI_HasUpdate = 2
    End If

    ie.Quit
    Set ie = Nothing

errHandler:
    If Err.Number <> 0 Then
        Debug.Print "Tulip Cell couldn't check for updates. " & Err.Description
    End If

done:
End Sub


Public Function TI_CallByName(name As String, ParamArray params() As Variant)
    On Error GoTo errHandler

    If (TI_HasUpdate = 1) Then
        TI_HasUpdate = 3
        MsgBox "There is a new version of Tulip Cell available." & vbCrLf & "Please visit https://tulipcell.org/ to update today.", vbInformation, "Tulip Cell"
    End If

    ChDir (ThisWorkbook.Path)

    Dim index As Long
    index = TI_GetIndicator(name)
    If (index < 0) Then
        MsgBox "Error. Couldn't find indicator index for " & name & "."
        GoTo done
    End If

    Dim input_count As Long, option_count As Long, output_count As Long
    input_count = TI_GetInputCount(index)
    option_count = TI_GetOptionCount(index)
    output_count = TI_GetOutputCount(index)

    If (UBound(params) + 1 <> input_count + option_count) Then
        MsgBox "Error: Wrong number of inputs or options for TI_CallByName(" & name & ")."
        GoTo done
    End If


    Dim size As Long
    size = params(0).Count

    Dim in_arr() As Double
    Dim opt_arr() As Double
    Dim out_arr() As Double

    ReDim in_arr(size * input_count)
    ReDim opt_arr(option_count)
    ReDim out_arr(size * output_count)


    Dim i As Long
    Dim pi As Long
    Dim cell As Variant
    i = 0
    For pi = 0 To input_count - 1
        If (params(pi).Count <> size) Then
            MsgBox "Error: All inputs are expected to be the same size."
            GoTo done
        End If

        For Each cell In params(pi)
            in_arr(i) = cell.Value
            i = i + 1
        Next cell
    Next pi


    For i = 0 To option_count - 1
        opt_arr(i) = params(i + input_count)
    Next i



    Dim ret As Long
    ret = TI_Call(index, size, in_arr(0), opt_arr(0), out_arr(0))

    If (ret <> 0) Then
        TI_CallByName = 0
        GoTo done
    End If

    Dim start As Long
    start = TI_GetStart(index, opt_arr(0))

    Dim out_shape() As Variant
    Dim col As Long, row As Long
    ReDim out_shape(size, output_count)
    For i = 0 To UBound(out_arr)
        col = Int(i / size)
        row = i Mod size
        If (row < start) then
            out_shape(row, col) = ""
        Else
            out_shape(row, col) = out_arr(i)
        End If
    Next i
    TI_CallByName = out_shape


errHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, Err.Number
    End If

done:

End Function
}



set indicators [exec $sample_path --list]
set indicators [split $indicators "\n"]

set register {}

set php_array {}
foreach i $indicators {
    array set ind $i

    set idx 0
    set inputs_dec {}
    set inputs_bare {}
    foreach i $ind(inputs) {
        if {$i eq "real"} {set i "input[incr idx]"}
        set i [string totitle $i]
        lappend inputs_dec "[set i]Range As Range"
        lappend inputs_bare [set i]Range
    }

    set options_dec {}
    set options_bare {}
    foreach o $ind(options) {
        set k {}
        foreach w $o {
            append k [string totitle $w]
        }
        set o [string map [list "%" ""] $k]
        lappend options_dec "$o As Double"
        lappend options_bare $o
    }

    puts $out "\n'$ind(full_name)"
    puts $out "Public Function TI_[string toupper $ind(name)]([join [concat $inputs_dec $options_dec] {, }])"
    puts $out "    TI_[string toupper $ind(name)] = TI_CallByName(\"$ind(name)\", [join [concat $inputs_bare $options_bare] {, }])"
    puts $out "End Function"

    lappend register "    Application.MacroOptions Macro:=\"TI_[string toupper $ind(name)]\", Description:=\"$ind(full_name)\", Category:=\"Tulip Cell Technical Analysis\""

}

puts $out "\n"
puts $out "Public Sub TI_RegisterHelp()"
puts $out "    On Error Resume Next 'Older Excel versions don't support the following functions"
puts $out [join $register \n]
puts $out "End Sub"




close $out
