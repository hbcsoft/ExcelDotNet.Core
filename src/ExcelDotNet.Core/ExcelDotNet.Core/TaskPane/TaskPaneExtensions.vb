Imports System.Runtime.CompilerServices
Imports Forms = System.Windows.Forms
Imports WPF = System.Windows.Controls


Namespace TaskPane


    Public Module TaskPaneExtensions

        ''' <summary>
        ''' Show the current user control in the task pane.
        ''' </summary>
        ''' <param name="userControl">The <see cref="WPF.UserControl"/> to show in the task pane.</param>
        ''' <param name="title">The title to use in the task pane.</param>
        ''' <returns></returns>
        <Extension>
        Public Function ShowInTaskPane(ByVal userControl As WPF.UserControl, ByVal title As String) As WpfTaskPane

            'Create and return the WpfTaskPane
            Return New WpfTaskPane(title:=title) With {.Visible = True, .UserControl = userControl}
        End Function

    End Module

End Namespace


