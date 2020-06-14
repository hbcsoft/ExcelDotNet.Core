Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports ExcelDna.Integration.CustomUI

Namespace TaskPane

    ''' <summary>
    ''' A ComVisible <see cref="UserControl"/> allowing the use with the <see cref="CustomTaskPane"/>.
    ''' </summary>
    ''' <remarks></remarks>
    <ComVisible(True)>
    <DesignerCategory("Code")>
    Public NotInheritable Class TaskPaneUserControl : Inherits UserControl

        Friend Sub New()
        End Sub
    End Class

End Namespace
