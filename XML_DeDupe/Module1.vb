Module Module1

    Dim assembly As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
    Dim fvi As FileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location)
    Dim companyName As String = fvi.CompanyName
    Dim productName As String = fvi.ProductName
    Dim productVersion As String = fvi.ProductVersion

    Dim strFilename As String

    Sub Main()
        Console.WriteLine()
        Dim args() As String = Environment.GetCommandLineArgs
        If args.Count > 1 Then
            strFilename = args(1)

            If IO.File.Exists(strFilename) Then
                Dim fi As IO.FileInfo = New IO.FileInfo(strFilename)
                If fi.Extension.ToLower <> ".xml" Then
                    Console.Write("This file does not have a "".xml"" extension.  Proceed? [y/n]:  ")
                    If Console.ReadKey.Key = ConsoleKey.Y Then
                        Console.WriteLine()
                        ProcessFile()
                    Else
                        Console.WriteLine()
                    End If
                Else
                    ProcessFile()
                End If
            Else
                Console.WriteLine()
                Console.WriteLine("File not found:  " & strFilename)
            End If


        Else
            WriteUsage()
        End If
    End Sub

    Private Sub ProcessFile()
        Console.WriteLine("Processing:  " & strFilename)
        Dim intDupes As Integer = 0
        Dim intRowsBefore As Integer = 0
        Dim intRowsAfter As Integer = 0

        Dim ds As New DataSet
        ds.ReadXml(strFilename)
        ds.AcceptChanges()

        Dim dt As DataTable = ds.Tables(0)
        intRowsBefore = dt.Rows.Count

        Dim columnNames() As String = dt.Columns.Cast(Of DataColumn)().[Select](Function(x) x.ColumnName).ToArray()
        dt = dt.DefaultView.ToTable(True, columnNames)
        intRowsAfter = dt.Rows.Count

        intDupes = intRowsBefore - intRowsAfter
        Console.WriteLine("Rows Before:" & vbTab & intRowsBefore)
        Console.WriteLine("Rows After:" & vbTab & intRowsAfter)

        If intDupes > 0 Then
            dt.WriteXml(strFilename)
        End If

        Console.WriteLine("Dupes:" & vbTab & intDupes)
        Console.WriteLine()

    End Sub

    Private Sub WriteUsage()
        Console.WriteLine()
        Console.WriteLine("Description:  Eliminate duplicates in an XML file")
        Console.WriteLine("Author:       Joe Ostrander")
        Console.WriteLine("Date:         2014.12.09")
        Console.WriteLine()
        Console.WriteLine("USAGE:")
        Console.WriteLine(productName & " <path to XML file>")
        Console.WriteLine()
        Console.WriteLine("Example:")
        Console.WriteLine(productName & " ""C:\Temp\My XML File.xml""")
        Environment.Exit(0)
    End Sub
End Module
