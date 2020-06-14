Imports System
Imports System.Runtime.CompilerServices
Imports Excel = NetOffice.ExcelApi

Namespace Extensions

    Public Module TableExtensions

        ''' <summary>
        ''' Gets the table associated with the specified name. Search is case insensitive.
        ''' </summary>
        ''' <param name="book"></param>
        ''' <param name="name">The name of the table to get.</param>
        ''' <returns>If the name if found, returns the table with the specified name; otherwise, return nothing."/></returns>
        <Extension()>
        Public Function GetTable(ByVal book As Excel.Workbook, ByVal name As String) As Excel.ListObject

            'Stores the object
            Dim tb As Excel.ListObject = Nothing

            'Check if the table exists all the worksheets
            For Each sh As Excel.Worksheet In book.Worksheets

                'Exit when found
                If TryGetTable(sheet:=sh, name:=name, table:=tb) Then
                    Exit For
                End If
            Next

            'Return
            Return tb
        End Function


        ''' <summary>
        ''' Gets the table associated with the specified name. Search is case insensitive.
        ''' </summary>
        ''' <param name="sheet"></param>
        ''' <param name="name">The name of the table to get.</param>
        ''' <param name="table">When this method returns, contains the table with the specified name, if the name is found; otherwise, nothing. This parameter is passed uninitialized.</param>
        ''' <returns>true if the <see cref="Excel.Worksheet"/> contains a table with the specified name; otherwise, false.</returns>
        <Extension()>
        Public Function TryGetTable(ByVal sheet As Excel.Worksheet, ByVal name As String, ByVal table As Excel.ListObject) As Boolean

            'Check all the objects in the worksheet and return the listObject if found
            'Search is case insensitive
            For Each tb As Excel.ListObject In sheet.ListObjects

                If String.Equals(tb.Name, name, StringComparison.OrdinalIgnoreCase) Then
                    table = tb
                    Return True
                End If
            Next

            'Return false if not found
            Return False
        End Function


    End Module

End Namespace
