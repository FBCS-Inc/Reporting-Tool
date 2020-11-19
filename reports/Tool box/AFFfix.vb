Imports System.IO
Imports System.Text.RegularExpressions
Public Class AFFfix
    Private kpe(4) As String
	Private delim As String               '"originator":{"servicer":"886d8f5f-0f0b-446b-b1f1-b3c2e1c72a0d"}
	Private findthis As String = Chr(34) + "originator" + Chr(34) + ":{" + Chr(34) + "servicer" + Chr(34) + ":" + Chr(34) + "886d8f5f-0f0b-446b-b1f1-b3c2e1c72a0d" + Chr(34) + "}"
	Private rplcthis As String = Chr(34) + "originator" + Chr(34) + ":" + Chr(34) + "creditor" + Chr(34)
    Private outfile As String
    Dim revised As String
    Dim dt1 As New DataTable
    Dim dt2 As New DataTable

    Private Sub AFFfix_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        kpe(0) = Chr(34) + "event" + Chr(34) + ":{" + Chr(34) + "tagadded" + Chr(34)
        kpe(1) = Chr(34) + "event" + Chr(34) + ":{" + Chr(34) + "recalled" + Chr(34)
        kpe(2) = Chr(34) + "event" + Chr(34) + ":{" + Chr(34) + "appliedFinancialTransaction" + Chr(34)
        kpe(3) = Chr(34) + "event" + Chr(34) + ":{" + Chr(34) + "ceaseAndDesistFiled" + Chr(34)
        delim = "{" + Chr(34) + "outSourcerAccountId" + Chr(34)
        dt1.Columns.Add("Line")
        dt1.Columns.Add("Data")
        dt2.Columns.Add("Line")
        dt2.Columns.Add("Data")
        DataGridView1.DataSource = dt1
    End Sub

    Private Sub Jsoninput_Click(sender As Object, e As EventArgs) Handles jsoninput.Click
        With OpenFileDialog1
            .FileName = jsonfile.Text
            .DefaultExt = "json"
            .InitialDirectory = "\\fbpa1fs1\DataEntry\New Business\AFF"
            .Filter = "json files (*.json)|*.json|json Files (*.json)|*.json"
            .ShowDialog()
        End With
    End Sub
    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        jsonfile.Text = OpenFileDialog1.FileName
        revised = OpenFileDialog1.FileName
        outfile = OpenFileDialog1.SafeFileName
        revised = revised.Replace(outfile, String.Empty)
		outfile = revised + "Revised\" + outfile
		Loadrawfile()
    End Sub

	Private Sub Openjson_Click(sender As Object, e As EventArgs)
		Loadrawfile()
	End Sub
	Private Sub Loadrawfile()
        dt1.Clear()
        Dim fs As FileStream
        Dim reader As StreamReader
        Dim Line As String = ""
        Dim newline As String = ""
        Dim jsonline As String()
        fs = New FileStream(jsonfile.Text, FileMode.Open)
        reader = New StreamReader(fs)
        Cursor = Cursors.WaitCursor
        Do

            Line = reader.ReadLine

            If Line Is Nothing Then Exit Do
            newline = Line.Replace(delim, Environment.NewLine + delim)
            jsonline = newline.Split(Environment.NewLine)

            For x As Integer = 0 To UBound(jsonline)
                dt1.Rows.Add(x.ToString, jsonline(x).ToString)
            Next


        Loop
        rawcount.Text = (dt1.Rows.Count - 1).ToString
        reader.Dispose()
        fs.Dispose()
        Cursor = Cursors.Default

    End Sub

    Private Sub Clean_Click(sender As Object, e As EventArgs) Handles Clean.Click
        LoadClean()
    End Sub
    Private Sub LoadClean()
        dt2.Clear()
        Dim rowstring As String
        Cursor = Cursors.WaitCursor
        For x As Integer = 1 To rawcount.Text
            rowstring = DataGridView1.Rows(x).Cells(1).Value

            For y As Integer = 0 To 3
                If rowstring.ToUpper.Contains(kpe(y).ToString.ToUpper) Then
					rowstring = rowstring.Replace(findthis, rplcthis)
					'rowstring = rowstring.Remove(InStr(rplcthis, rowstring) + Len(rplcthis) + 1, 37)
					dt2.Rows.Add(x.ToString, rowstring)
                    Exit For
                End If
            Next y
        Next x
        DataGridView2.DataSource = dt2
        cleancount.Text = dt2.Rows.Count.ToString
        Cursor = Cursors.Default
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Saveoutfile()
    End Sub
	Private Sub Saveoutfile()
		Cursor = Cursors.WaitCursor
		Dim jsonstring As String = "["
		For x As Integer = 0 To dt2.Rows.Count - 1
			jsonstring += dt2.Rows(x).Item(1).ToString
		Next
		jsonstring = jsonstring.Replace(vbLf, [String].Empty)
		jsonstring = jsonstring.Replace(vbCr, [String].Empty)
		jsonstring = jsonstring.Replace(vbTab, [String].Empty)
		jsonstring = jsonstring.Substring(0, jsonstring.Length - 1) + "]"
		Dim File As System.IO.StreamWriter
		If (Not System.IO.Directory.Exists(revised + "Revised\")) Then
			System.IO.Directory.CreateDirectory(revised + "Revised\")
		End If
		File = My.Computer.FileSystem.OpenTextFileWriter(outfile, True)
		File.Write(jsonstring)
		File.Close()
		Cursor = Cursors.Default
		MessageBox.Show("Cleaned input file: " + outfile + " is complete.  You can process this file", "Clean up Complete")
	End Sub
End Class