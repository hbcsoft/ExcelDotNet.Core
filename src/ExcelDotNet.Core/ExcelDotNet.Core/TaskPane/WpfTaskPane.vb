Imports System.Windows.Forms.Integration
Imports Dna = ExcelDna.Integration.CustomUI
Imports Forms = System.Windows.Forms
Imports WPF = System.Windows.Controls

Namespace TaskPane

    ''' <summary>
    ''' An object for manipulating Excel Task Panes that contains Windows Presentation Foundation controls. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class WpfTaskPane

        ''' <summary>
        ''' Creates a new Excel task paea that support Wpf controls.
        ''' </summary>
        ''' <param name="title"></param>
        Public Sub New(ByVal title As String)

            'Controls only can be used in CustomTaskPane if -Is ComVisible and -Is type of System.Windows.Forms.UserControl
            'So, the BaseTaskPaneUsedControl is used
            fUserControl = New TaskPaneUserControl

            'Creates and store the CustomTaskPane
            CustomTaskPane = Dna.CustomTaskPaneFactory.CreateCustomTaskPane(userControl:=fUserControl, title:=title)

            'Initializes a new HostElement and sets the basic properties
            'Needs to be invisible until has the control inside
            elementHost = New ElementHost With {.Dock = Forms.DockStyle.Fill, .Visible = False}


            'Adds the element host to the base control
            fUserControl.Controls.Add(elementHost)

        End Sub


        ''' <summary>
        ''' Stores the windows form user control that is used in the task pane.
        ''' </summary>
        Private fUserControl As Forms.UserControl


        ''' <summary>
        ''' Store the windows Forms Element Host that prentent the WPF control.
        ''' </summary>
        Private elementHost As ElementHost



        ''' <summary>
        ''' Gets the <see cref="Dna.CustomTaskPane"/> Interface.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property CustomTaskPane As Dna.CustomTaskPane

        ''' <summary>
        ''' Gets or sets the visibility of the Task pane.
        ''' </summary>
        ''' <returns></returns>
        Public Property Visible As Boolean
            Get
                Return CustomTaskPane.Visible
            End Get
            Set(value As Boolean)
                CustomTaskPane.Visible = value
            End Set
        End Property


        Private _UserControl As WPF.UserControl
        ''' <summary>
        ''' Gets or sets the <see cref="WPF.UserControl"/> that is presented in the task pane.
        ''' </summary>
        ''' <value></value>
        Public Property UserControl() As WPF.UserControl
            Get
                Return _UserControl
            End Get
            Set(ByVal value As WPF.UserControl)
                'The element host needs to be invisible when child control is set
                elementHost.Visible = False
                _UserControl = value
                elementHost.Child = value
                elementHost.Visible = True
            End Set
        End Property

    End Class

End Namespace

