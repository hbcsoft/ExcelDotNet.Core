Imports ExcelDna.Integration
Imports ExcelDotNet.Core.TaskPane

Public Module ExcelCommands

    <ExcelCommand>
    Public Sub OpenTaskPane()

        Dim uc As New UserControl1
        wtp = uc.ShowInTaskPane("My control")

    End Sub


    Private wtp As WpfTaskPane


End Module
