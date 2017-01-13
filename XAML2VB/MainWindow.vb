Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Markup
Imports System.Xml
<Assembly: AssemblyTitle("Precompilador XAML")>
<Assembly: AssemblyDescription("Precompilador de archivos XAML")>
<Assembly: AssemblyCompany("TheXDS! non-Corp.")>
<Assembly: AssemblyProduct("MCART")>
<Assembly: AssemblyCopyright("Copyright © 2017 TheXDS! non-Corp.")>
<Assembly: AssemblyTrademark("MCA Corp.")>
<Assembly: ComVisible(False)>
<Assembly: ThemeInfo(ResourceDictionaryLocation.None, ResourceDictionaryLocation.SourceAssembly)>
<Assembly: Guid("59afe06b-508e-4a3a-8807-80f8a2b45c6e")>
<Assembly: AssemblyVersion("1.0.0.0")>
<Assembly: AssemblyFileVersion("1.0.0.0")>
Public MustInherit Class XAMLPreCompiler
    Protected Domain As New List(Of Assembly)
    Protected Const DefaultNameSpace As String = "#"
    Protected [Imports] As New StringBuilder
    Protected importedNames As New Dictionary(Of String, String)
    Protected Declarations As New StringBuilder
    Protected NewDeclrs As New StringBuilder
    Protected NewMethod As New StringBuilder
    Protected ErrOutput As New StringBuilder
    Protected ModuleGen As New StringBuilder
    Protected xamlreader As New XmlDocument
    Protected InstanceCounter As New Dictionary(Of String, Integer)
    Protected MustOverride ReadOnly Property FileExtension As String
    Protected MustOverride Function IsComment(x As String) As Boolean
    Protected MustOverride ReadOnly Property CommentFormat As String
    Public MustOverride Function Convert(XAMLFile As String, Optional AdditionalAssemblies As List(Of Assembly) = Nothing) As String
    Protected MustOverride Function ProcessElement(x As XmlElement, Optional iselement As Boolean = True, Optional parent As Object = Nothing) As String
    Protected MustOverride Function GetExpression(Expr As String, Optional SuggestedType As Type = Nothing, Optional PrName As String = Nothing) As String
    Public MustOverride ReadOnly Property DisplayName As String
    Public MustOverride ReadOnly Property Description As String
    Protected Function InferType(PrName As String) As Type
        Select Case PrName
            Case "Content", "Source"
                Return GetType(Object)
            Case "Foreground", "Background"
                Return GetType(Brush)
            Case "Margin"
                Return GetType(Thickness)
            Case "Height", "Width", "Top",
                 "Left", "Lenght", "FontSize",
                 "Size", "Minimum", "Maximum",
                 "MinLength", "MaxLength", "Radius",
                 "ButtonWidth", "ButtonHeight",
                 "MinWidth", "MinHeight", "MaxWidth",
                 "MaxHeight", "Grid.Row", "Grid.Column",
                 "Grid.RowSpan", "Grid.ColumnSpan"
                Return GetType(Double)
            Case "NotifyOnValidationError",
                 "ValidatesOnExceptions",
                 "IsReadOnly", "IsEnabled",
                 "IsChecked", "IsEmpty", "IsVisible"
                Return GetType(Boolean)
            Case "SizeToContent", "ResizeMode", "Icon",
                 "Visibility", "HorizontalAlignment",
                 "VerticalAlignment", "TextWrapping"
                Return FindType(PrName)
            Case "DockPanel.Dock"
                Return GetType(Dock)
        End Select
        Return GetType(String)
    End Function
    Protected Function FindType(ByVal name As String) As Type
        For Each asmbly As Assembly In Domain
            For Each j As TypeInfo In asmbly.DefinedTypes
                If (name.Contains(".") AndAlso j.FullName = name) OrElse j.Name = name Then
                    Return j.AsType
                End If
            Next
        Next
        Return Nothing
    End Function
    Protected Function ParseType(ByVal ts As String) As Type
        Dim ns As String = Nothing
        If ts.Contains(":") Then
            ns = importedNames(ts.Substring(0, ts.IndexOf(":"))) & "."
            ts = ts.Substring(ts.IndexOf(":") + 1)
        End If
        If ts.Contains(".") Then
            Return FindType(ns & ts.Substring(0, ts.IndexOf("."))) _
                .GetProperty(ts.Substring(ts.IndexOf(".") + 1)).PropertyType
        Else
            Return FindType(ns & ts)
        End If
    End Function
    Protected Function Splt(s As String, Optional sepchar As Char = ",") As String()
        Dim l As New List(Of String)
        Dim i As String = ""
        Dim subp As Integer = 0
        For Each j As Char In s.ToCharArray
            Select Case j
                Case "{"
                    subp += 1
                    i &= "{"
                Case "}"
                    subp -= 1
                    i &= "}"
                Case sepchar
                    If subp = 0 Then
                        l.Add(i.Trim)
                        i = ""
                    Else
                        i &= sepchar
                    End If
                Case Else
                    i &= j
            End Select
        Next
        If i <> "" Then l.Add(i.Trim)
        Return l.ToArray
    End Function
    Protected Function Clip(s As String) As String
        Return s.Substring(1, s.Length - 2)
    End Function
    Protected Sub ErrComment(x As String)
        For Each j As String In x.Split(vbCrLf)
            ErrOutput.AppendFormat(CommentFormat, j)
            ErrOutput.AppendLine()
        Next
    End Sub
End Class
Public Class XAMLVBPreCompiler
    Inherits XAMLPreCompiler
    Public Overrides Function Convert(XAMLFile As String, Optional AdditionalAssemblies As List(Of Assembly) = Nothing) As String
        Dim vbfile As StreamReader
        Dim outpf As StreamWriter
        [Imports] = New StringBuilder
        importedNames = New Dictionary(Of String, String)
        Declarations = New StringBuilder
        NewDeclrs = New StringBuilder
        NewMethod = New StringBuilder
        ErrOutput = New StringBuilder
        XamlReader = New XmlDocument
        Domain.Clear()
        Try
            vbfile = New StreamReader(XAMLFile & "." & FileExtension)
        Catch ex As FileNotFoundException
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Return Nothing
        End Try
        Dim Fname As String = XAMLFile.Substring(0, XAMLFile.Length - 4) & FileExtension
        Dim ClassName As String = Nothing
        Dim sname As String = Nothing
        Dim sname2 As String = Nothing
        Dim asm As XmlnsDefinitionAttribute = Nothing
        XamlReader.Load(XAMLFile)
        Declarations.AppendLine("    Inherits " & XamlReader.DocumentElement.Name)
        For Each j As XmlAttribute In XamlReader.DocumentElement.Attributes
            If j.Name.StartsWith("xmlns") Then
                If j.Value.StartsWith("clr-namespace:") Then
                    If j.Value.Contains(";") Then
                        sname = j.Value.Substring(14, j.Value.IndexOf(";") - 14)
                        For Each k As Assembly In AppDomain.CurrentDomain.GetAssemblies
                            If k.GetName.Name = j.Value.Substring(j.Value.LastIndexOf("=") + 1) _
                                AndAlso Not Domain.Contains(k) Then Domain.Add(k)
                        Next
                    Else
                        sname = j.Value.Substring(14)
                    End If
                Else
                    sname = Nothing
                    For Each asmbly As Assembly In AppDomain.CurrentDomain.GetAssemblies
                        For Each a As XmlnsDefinitionAttribute In Attribute.GetCustomAttributes(asmbly, GetType(XmlnsDefinitionAttribute))
                            If a IsNot Nothing AndAlso a.XmlNamespace = j.Value AndAlso Not Domain.Contains(asmbly) Then
                                Domain.Add(asmbly)
                                [Imports].AppendLine("Imports " & a.ClrNamespace)
                            End If
                        Next
                    Next
                End If
                If j.Name.Contains(":") Then
                    sname2 = j.Name.Substring(6)
                Else
                    sname2 = DefaultNameSpace
                End If
                If sname <> Nothing Then
                    [Imports].AppendLine("Imports " & sname)
                    importedNames.Add(sname2, sname)
                End If
            ElseIf j.Name = "x:Class" Then
                ClassName = j.Value
            End If
        Next
        If AdditionalAssemblies IsNot Nothing Then
            For Each j As Assembly In AdditionalAssemblies
                Domain.Add(j)
            Next
        End If
        If ClassName = Nothing Then Throw New InvalidXAMLException
        ProcessElement(XamlReader.DocumentElement, False)
        outpf = New StreamWriter(Fname, FileMode.Create)
        outpf.Write([Imports].ToString)
        Dim lnebuff As String
        Dim newFound As Boolean
        Dim classfound As Boolean = False
        Do Until vbfile.EndOfStream
            lnebuff = vbfile.ReadLine
            If lnebuff.Trim.StartsWith("End Class") AndAlso classfound AndAlso Not newFound Then
                outpf.WriteLine("    ''' <summary>")
                outpf.WriteLine("    ''' Crea una nueva instancia de esta clase")
                outpf.WriteLine("    ''' </summary>")
                outpf.WriteLine("    Public Sub New()")
                outpf.Write(NewDeclrs.ToString)
                outpf.Write(NewMethod.ToString)
                outpf.WriteLine("    End Sub")
            End If
            outpf.WriteLine(lnebuff)
            lnebuff = " " & lnebuff & " "
            If lnebuff.Contains(" Class " & ClassName & " ") Then
                classfound = True
                outpf.Write(Declarations.ToString)
            End If
            If lnebuff.Contains(" Sub New(") Then
                newFound = True
                outpf.Write(NewDeclrs.ToString)
                outpf.Write(NewMethod.ToString)
            End If
        Loop
        If ModuleGen.ToString <> "" Then
            outpf.WriteLine(ModuleGen.ToString)
        End If
        outpf.Write(ErrOutput.ToString)
        outpf.Flush()
        outpf.Close()
        vbfile.Close()
        Domain = Nothing
        Return Fname
    End Function
    Protected Overrides ReadOnly Property FileExtension As String
        Get
            Return "vb"
        End Get
    End Property
    Public Overrides ReadOnly Property DisplayName As String
        Get
            Return "Precompilador a VB.Net"
        End Get
    End Property
    Public Overrides ReadOnly Property Description As String
        Get
            Return "Procesa un archivo XAML (con el XAML.vb que lo acompaña) y crea una nueva clase compilable incrustando los elementos XAML como objetos de Visual Basic."
        End Get
    End Property
    Protected Overrides ReadOnly Property CommentFormat As String
        Get
            Return "' {0}"
        End Get
    End Property
    Protected Overrides Function ProcessElement(x As XmlElement, Optional iselement As Boolean = True, Optional parent As Object = Nothing) As String
        Dim t As Type = ParseType(x.Name)
        Dim v As XmlAttribute = Nothing
        Dim nme As String = String.Empty
        Dim withadded As Boolean = False
        Dim StrBuilder As StringBuilder = Nothing
        If t Is Nothing Then Throw New SyntaxError
        v = x.GetAttributeNode("x:Name")
        If iselement Then
            If v Is Nothing Then
                If InstanceCounter.ContainsKey(t.Name) Then
                    InstanceCounter.Item(t.Name) += 1
                Else
                    InstanceCounter.Add(t.Name, 1)
                End If
                nme = t.Name & InstanceCounter.Item(t.Name)
                StrBuilder = NewDeclrs
                StrBuilder.Append("        Dim")
            Else
                nme = v.Value
                StrBuilder = Declarations
                StrBuilder.Append("    Private WithEvents")
            End If
            StrBuilder.Append(String.Format(
                " {0} As {1}{2}", nme,
                If(t.IsClass, "New ", ""),
                t.FullName))
        Else
            If v IsNot Nothing Then
                nme = v.Value
                StrBuilder = ModuleGen
                StrBuilder.AppendLine("Public Module " & nme & "Definition")
                StrBuilder.Append("    Public WithEvents")
                StrBuilder.Append(String.Format(
                    " {0} As {1}{2}", nme,
                    If(t.IsClass, "New ", ""),
                    x.GetAttribute("x:Class")))
                iselement = True
            Else
                nme = "Me"
            End If
        End If
        For Each j As PropertyInfo In t.GetProperties
            v = x.GetAttributeNode(j.Name)
            If v IsNot Nothing Then
                If v.Value.StartsWith("{Binding") Then
                    NewMethod.AppendLine(String.Format(
                        "        {0}.SetBinding({1}.{2}Property, {3})",
                        nme, t.FullName, v.Name,
                        GetExpression(v.Value, j.PropertyType, j.Name)))
                ElseIf iselement AndAlso Not v.Value.Contains("{StaticResource ") Then
                    If Not withadded Then
                        StrBuilder.AppendLine(" With {")
                        withadded = True
                    Else
                        StrBuilder.AppendLine(",")
                    End If
                    StrBuilder.Append(Space(12))
                    StrBuilder.Append(String.Format(".{0} = {1}",
                        j.Name, GetExpression(v.Value, j.PropertyType, v.Name)))
                Else
                    NewMethod.AppendLine(String.Format(
                        "        {0}.{1} = {2}",
                        nme, v.Name, GetExpression(v.Value, j.PropertyType, v.Name)))
                End If
            End If
        Next
        If withadded Then StrBuilder.Append("}")
        If StrBuilder IsNot Nothing Then StrBuilder.AppendLine()
        If StrBuilder Is ModuleGen Then StrBuilder.AppendLine("End Module")
        For Each j As EventInfo In t.GetEvents
            v = x.GetAttributeNode(j.Name)
            If v IsNot Nothing Then
                NewMethod.AppendLine(String.Format(
                    "        AddHandler {0}.{1}, AddressOf {2}",
                    nme, v.Name, v.Value))
            End If
        Next
        For Each j As String In {"Grid.Row",
            "Grid.Column",
            "Grid.RowSpan",
            "Grid.ColumnSpan",
            "DockPanel.Dock",
            "FrameworkElement.FlowDirection",
            "Grid.IsSharedSizeScope",
            "TextBlock.BaselineOffset",
            "TextBlock.FontFamily",
            "TextBlock.FontSize",
            "TextBlock.FontStretch",
            "TextBlock.FontStyle",
            "TextBlock.FontWeight",
            "TextBlock.Foreground",
            "TextBlock.LineHeight",
            "TextBlock.LineStackingStrategy",
            "TextBlock.TextAlignment",
            "Panel.ZIndex"}
            v = x.GetAttributeNode(j)
            If v IsNot Nothing Then
                NewMethod.Append(Space(8))
                NewMethod.Append(v.Name.Split(".")(0) & ".Set" & v.Name.Split(".")(1) & "(" & nme & ", ")
                NewMethod.AppendLine(GetExpression(v.Value, , v.Name) & ")")
            End If
        Next
        For Each j As XmlNode In x.ChildNodes
            Select Case j.Name
                Case t.Name & ".Resources"
                    For Each k As XmlElement In j.ChildNodes
                        NewMethod.AppendLine(String.Format(
                        "        {0}.Resources.Add(""{1}"", {2})", nme,
                        k.GetAttribute("x:Key"),
                        ProcessElement(k)))
                    Next
                Case t.Name & ".RowDefinitions", t.Name & ".ColumnDefinitions"
                    For Each k As XmlElement In j.ChildNodes
                        NewMethod.AppendLine(String.Format(
                            "        {0}.{1}.Add({2})", nme,
                            j.Name.Substring(j.Name.IndexOf(".") + 1),
                            ProcessElement(k)))
                    Next
                Case Else
                    If j.Name.StartsWith(t.Name & ".") Then
                        NewMethod.AppendLine(String.Format(
                            "        {0}.{1} = {2}", nme,
                            j.Name.Substring(j.Name.IndexOf(".") + 1),
                            ProcessElement(j)))
                    ElseIf t.GetProperty("Content") IsNot Nothing Then
                        NewMethod.AppendLine(String.Format("        {0}.Content = {1}", nme, ProcessElement(j)))
                    ElseIf t.GetProperty("Children") IsNot Nothing Then
                        NewMethod.AppendLine(String.Format("        {0}.Children.Add({1})", nme, ProcessElement(j)))
                    ElseIf Not t.IsClass Then
                        NewMethod.AppendLine(String.Format("        {0} = {1}", nme, GetExpression(j.Value, t)))
                    End If
            End Select
        Next
        Return nme
    End Function
    Protected Overrides Function GetExpression(Expr As String, Optional SuggestedType As Type = Nothing, Optional PrName As String = Nothing) As String
        Dim args As String() = Nothing
        If Expr.ToCharArray.First = "{" Then
            Expr = Clip(Expr.Replace(vbCrLf, Nothing))
            args = Splt(Expr)
            Select Case args.First.Split(" ").First
                Case "Binding"
                    Dim p As String = Nothing
                    Dim t As New StringBuilder
                    Dim withadded As Boolean = False
                    For Each j As String In args
                        j = j.Trim
                        If j.Contains("=") Then
                            With Splt(j, "=")
                                If Not withadded Then
                                    t.AppendLine(" With {")
                                    t.Append(Space(12))
                                    withadded = True
                                Else
                                    t.AppendLine(",")
                                    t.Append(Space(12))
                                End If
                                If .ElementAt(0).Split.Last.Trim = "ElementName" Then
                                    t.Append(".Source = ")
                                    t.Append(GetExpression(.ElementAt(1), GetType(DependencyObject)))
                                Else
                                    t.Append("." & .ElementAt(0).Split.Last.Trim & " = ")
                                    t.Append(GetExpression(.ElementAt(1), GetType(Binding).GetProperty(
                                        .ElementAt(0).Split.Last.Trim).PropertyType))
                                End If
                            End With
                        Else
                            p = "(""" & j.Substring(8).Trim & """)"
                        End If
                    Next
                    If withadded Then t.Append("}")

                    Return "New Binding" & p & t.ToString
                Case "StaticResource"
                    Return "FindResource(""" & Split(args.First).Last & """)"
                Case "DynamicResource"
                    ' TODO: implementar DynamicResource
            End Select
        Else
            If SuggestedType = Nothing Then SuggestedType = InferType(PrName)
            Select Case SuggestedType
                Case GetType(String), GetType(Char), GetType(Object)
                    Return """" & Expr & """"
                Case GetType(Byte), GetType(SByte), GetType(Short), GetType(UShort),
                    GetType(Integer), GetType(UInteger), GetType(Long), GetType(ULong),
                    GetType(Single), GetType(Double), GetType(Decimal), GetType(IntPtr),
                    GetType(UIntPtr), GetType(Boolean), GetType(DependencyObject)
                    Return Expr
                Case GetType(Thickness)
                    args = Splt(Expr)
                    For j As Integer = 0 To args.Count - 1
                        If args(j).ToCharArray.First = "{" Then args(j) = GetExpression(args(j))
                    Next
                    Select Case args.Count
                        Case 1
                            Return "New Thickness(" & args(0) & ")"
                        Case 2
                            Return "New Thickness(" &
                            args(0) & ", " & args(1) & ", " &
                            args(0) & ", " & args(1) & ")"
                        Case 4
                            Return "New Thickness(" &
                            args(0) & ", " & args(1) & ", " &
                            args(2) & ", " & args(3) & ")"
                        Case Else
                            Throw New SyntaxError
                    End Select
                Case GetType(Point)
                    args = Splt(Expr)
                    For j As Integer = 0 To args.Count - 1
                        If args(j).ToCharArray.First = "{" Then args(j) = GetExpression(args(j))
                    Next
                    Select Case args.Count
                        Case 2
                            Return "New Point(" & args(0) & ", " & args(1) & ")"
                        Case Else
                            Throw New SyntaxError
                    End Select
                Case GetType(Brush)
                    If Expr.ToCharArray.First = "#" Then
                        Return String.Format(
                            "New SolidColorBrush(New Color With " &
                            "{{.A = {0}, .R = {1}, .G = {2}, .B = {3}}})",
                            "&H" & Expr.Substring(1, 2),
                            "&H" & Expr.Substring(7, 2),
                            "&H" & Expr.Substring(5, 2),
                            "&H" & Expr.Substring(3, 2))
                    Else
                        Return "Brushes." & Expr
                    End If
                Case GetType(Color)
                    If Expr.ToCharArray.First = "#" Then
                        Return String.Format(
                            "New Color With " &
                            "{{.A = {0}, .R = {1}, .G = {2}, .B = {3}}}",
                            "&H" & Expr.Substring(1, 2),
                            "&H" & Expr.Substring(7, 2),
                            "&H" & Expr.Substring(5, 2),
                            "&H" & Expr.Substring(3, 2))
                    Else
                        Return "Colors." & Expr
                    End If
                Case GetType(GridLength)
                    If Expr.ToCharArray.Last = "*" Then
                        Return String.Format("New GridLength({0}, GridUnitType.Star)", Val(Expr))
                    Else
                        Return String.Format("New GridLength({0}, GridUnitType.Pixel)", Val(Expr))
                    End If
                Case GetType(ImageSource)
                    Return String.Format("New BitmapImage(New Uri(""{0}""))", Expr)
                Case Else
                    If SuggestedType.IsEnum Then
                        Return SuggestedType.FullName.Replace("+", ".") & "." & Expr
                    Else
                        Throw New SyntaxError
                    End If
            End Select
        End If
        ErrComment("TODO: Unable to parse XAML expression:")
        ErrComment(Expr)
        Return "Nothing"
    End Function
    Protected Overrides Function IsComment(x As String) As Boolean
        Return x.Trim.StartsWith("'")
    End Function
End Class
Public Class SyntaxError
    Inherits Exception
End Class
Public Class SelConv
    Implements IValueConverter
    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Return value >= 0
    End Function
    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Return If(value, 0, -1)
    End Function
End Class
Public Class InvalidXAMLException
    Inherits Exception
End Class
Public Class MainWindow
    Inherits Window
    Private WithEvents BtnChoose As New Button With {
        .Content = "..."}
    Private WithEvents TxtFName As New TextBox
    Private WithEvents BtnGo As New Button With {
        .Content = "Procesar",
        .Width = 60}
    Private WithEvents cmbConvs As New ComboBox With {
        .Margin = New Thickness(10, 0, 10, 0),
        .VerticalAlignment = VerticalAlignment.Center}
    Private WithEvents BtnAdd As New Button With {
        .Content = "+"}
    Private WithEvents BtnRem As New Button With {
        .Content = "-"}
    Private WithEvents LstAsm As New ListBox With {
        .SelectionMode = SelectionMode.Single}
    Private ofd As New Microsoft.Win32.OpenFileDialog With {
        .Filter = "Archivos XAML|*.xaml|Todos los archivos|*.*",
        .Multiselect = False,
        .CheckFileExists = True,
        .CheckPathExists = True}
    Private ofdasm As New Microsoft.Win32.OpenFileDialog With {
        .Filter = "Ensablados .NET|*.dll;*.exe|Todos los archivos|*.*",
        .Multiselect = True,
        .CheckFileExists = True,
        .CheckPathExists = True}
    Private asmbs As New List(Of Assembly)
    Private precs As New List(Of XAMLPreCompiler)
    Private Sub wnd1_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        For Each j As Assembly In AppDomain.CurrentDomain.GetAssemblies
            For Each k As Type In j.GetTypes
                If (Not k.IsAbstract) AndAlso GetType(XAMLPreCompiler).IsAssignableFrom(k) Then
                    precs.Add(k.GetConstructor(Type.EmptyTypes).Invoke({}))
                    cmbConvs.Items.Add(precs.Last.DisplayName)
                    cmbConvs.SelectedIndex = 0
                End If
            Next
        Next
    End Sub
    Private Sub BtnChoose_Click(sender As Object, e As RoutedEventArgs) Handles BtnChoose.Click
        With ofd
            If .ShowDialog Then
                TxtFName.Text = .FileName
            End If
        End With
    End Sub
    Private Sub BtnAdd_Click(sender As Object, e As RoutedEventArgs) Handles BtnAdd.Click
        With ofdasm
            If .ShowDialog Then
                Dim x As New Text.StringBuilder
                For Each j As String In .FileNames
                    Try
                        Dim asm As Assembly = Assembly.LoadFile(j)
                        asmbs.Add(asm)
                        LstAsm.Items.Add(asm.FullName)
                    Catch ex As Exception
                        x.AppendLine(ex.Message)
                    End Try
                Next
                If x.ToString <> "" Then MessageBox.Show(x.ToString, "Advertencia", MessageBoxButton.OK, MessageBoxImage.Warning)
            End If
        End With
    End Sub
    Private Sub BtnRem_Click(sender As Object, e As RoutedEventArgs) Handles BtnRem.Click
        asmbs.RemoveAt(LstAsm.SelectedIndex)
        LstAsm.Items.RemoveAt(LstAsm.SelectedIndex)
        LstAsm.SelectedIndex = -1
    End Sub
    Private Sub BtnGo_Click(sender As Object, e As RoutedEventArgs) Handles BtnGo.Click
        Try
            precs.ElementAt(cmbConvs.SelectedIndex).Convert(TxtFName.Text, asmbs)
            MessageBox.Show("Operación finalizada correctamente.", Title, MessageBoxButton.OK, MessageBoxImage.Information)
        Catch ex As SyntaxError
            MessageBox.Show("No se pudo realizar la conversión (¿Agregó las referencias a ensamblados externos?)",
                            "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        Catch ex As IO.FileNotFoundException
            MessageBox.Show("No se encontró el archivo " & TxtFName.Text,
                            "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
    Private Sub cmbConvs_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbConvs.SelectionChanged
        If cmbConvs.SelectedIndex >= 0 Then
            cmbConvs.ToolTip = New ToolTip With {.Content = precs.ElementAt(cmbConvs.SelectedIndex).Description}
        Else
            cmbConvs.ToolTip = Nothing
        End If
    End Sub
    ''' <summary>
    ''' Crea una nueva instancia de esta clase
    ''' </summary>
    Public Sub New()
        Dim SelConv1 As New SelConv
        Dim Double1 As Double
        Dim Thickness1 As Thickness
        Dim DockPanel1 As New DockPanel
        Dim DockPanel2 As New DockPanel
        Dim TextBlock1 As New TextBlock With {
            .Text = "Archivo de origen: "}
        Dim DockPanel3 As New DockPanel
        Dim TextBlock2 As New TextBlock With {
            .Text = "Precompilador:",
            .VerticalAlignment = VerticalAlignment.Center}
        Dim DockPanel4 As New DockPanel
        Dim TextBlock3 As New TextBlock With {
            .Text = "Ensamblados referenciados"}
        Dim DockPanel5 As New DockPanel
        Dim Grid1 As New Grid
        Dim RowDefinition1 As New RowDefinition
        Dim RowDefinition2 As New RowDefinition
        Title = "Precompilador XAML"
        Width = 400
        Height = 250
        Resources.Add("scnv", SelConv1)
        Double1 = 20
        Resources.Add("btnwidth", Double1)
        Thickness1 = New Thickness(10)
        Resources.Add("pnlmargin", Thickness1)
        DockPanel2.Margin = FindResource("pnlmargin")
        DockPanel.SetDock(DockPanel2, Dock.Top)
        DockPanel.SetDock(TextBlock1, Dock.Left)
        DockPanel2.Children.Add(TextBlock1)
        BtnChoose.Width = FindResource("btnwidth")
        DockPanel.SetDock(BtnChoose, Dock.Right)
        DockPanel2.Children.Add(BtnChoose)
        DockPanel2.Children.Add(TxtFName)
        DockPanel1.Children.Add(DockPanel2)
        DockPanel3.Margin = FindResource("pnlmargin")
        DockPanel.SetDock(DockPanel3, Dock.Bottom)
        DockPanel.SetDock(TextBlock2, Dock.Left)
        DockPanel3.Children.Add(TextBlock2)
        BtnGo.SetBinding(IsEnabledProperty, New Binding("SelectedIndex") With {
            .Source = cmbConvs,
            .Converter = FindResource("scnv")})
        DockPanel.SetDock(BtnGo, Dock.Right)
        DockPanel3.Children.Add(BtnGo)
        DockPanel3.Children.Add(cmbConvs)
        DockPanel1.Children.Add(DockPanel3)
        DockPanel4.Margin = FindResource("pnlmargin")
        DockPanel.SetDock(TextBlock3, Dock.Top)
        DockPanel4.Children.Add(TextBlock3)
        DockPanel.SetDock(Grid1, Dock.Right)
        Grid1.RowDefinitions.Add(RowDefinition1)
        Grid1.RowDefinitions.Add(RowDefinition2)
        BtnAdd.Width = FindResource("btnwidth")
        Grid1.Children.Add(BtnAdd)
        BtnRem.Width = FindResource("btnwidth")
        BtnRem.SetBinding(IsEnabledProperty, New Binding("SelectedIndex") With {
            .Source = LstAsm,
            .Converter = FindResource("scnv")})
        Grid.SetRow(BtnRem, 1)
        Grid1.Children.Add(BtnRem)
        DockPanel5.Children.Add(Grid1)
        DockPanel5.Children.Add(LstAsm)
        DockPanel4.Children.Add(DockPanel5)
        DockPanel1.Children.Add(DockPanel4)
        Content = DockPanel1
    End Sub
End Class
Public Module MainModule
    Sub Main()
        With New MainWindow
            .ShowDialog()
        End With
    End Sub
End Module