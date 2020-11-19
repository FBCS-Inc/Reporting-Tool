Imports System.Data.SqlClient

Public Class PPACreditAudit

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With ReportData1
            .RPTNAME = "PPA Credit Audit"
            .FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
            .shfmt = "TDTTDTTT"
            .vsql = "Select DISTINCT master.number AS [File Number], Convert(varchar, master.lastpaid, 101) As [Last Paid], master.desk, ISNULL(ppacreditdesk.TheData, '') AS PPACreditDesk, ISNULL(PPACreditDate.TheData, '') AS PPACreditDate, "
            .vsql += " ISNULL(Users.UserName, '') AS [Users Name], CASE WHEN ppacreditdesk.TheData IS NULL OR ppacreditdesk.TheData Not Like 'A0%' THEN '' WHEN [Users].UserName IS NOT NULL AND EXISTS "
            .vsql += " (SELECT * From notes n WITH (nolock) "
            .vsql += " Where n.number = master.number And n.user0 = users.loginname And n.comment Like 'Misc Extra%ppa%term%') THEN 'Y' ELSE 'N' END AS [PPA Terms Added?], CASE WHEN ppacreditdesk.TheData IS NULL OR "
            .vsql += " ppacreditdesk.TheData Not Like 'A0%' THEN '' WHEN (EXISTS (SELECT *  From notes n WITH (nolock)  "
            .vsql += " Where n.number = master.number And n.user0 = users.loginname And n.comment Like '%part2%' AND n.action = 'PROM' AND n.result = 'ADD') OR"
            .vsql += " ((Select COUNT(*) From PDC_View_FBCS p WITH (nolock) Where p.number = master.number And p.CreatedBy = users.loginname) + "
            .vsql += " (SELECT COUNT(*) From PDCC_View_FBCS p WITH (nolock) Where p.number = master.number And p.CreatedBy = users.loginname) + "
            .vsql += " (SELECT COUNT(*) From Promises_View_FBCS p WITH (nolock) Where p.AcctID = master.number And p.CreatedBy = users.loginname)) >= 2) Then 'Y' ELSE 'N' END AS [2+ Arrangements Added?] "
            .vsql += " From master WITH (nolock) "
            .vsql += " INNER Join desk WITH (nolock) ON desk.code = master.desk "
            .vsql += " LEFT OUTER Join MiscExtra As ppacreditdesk WITH (nolock) ON ppacreditdesk.Number = master.number And ppacreditdesk.Title Like 'ppa%credit%desk%' "
            .vsql += " LEFT OUTER JOIN MiscExtra AS PPACreditDate WITH (nolock) ON PPACreditDate.Number = master.number And PPACreditDate.Title Like 'PPACreditDate' "
            .vsql += " LEFT OUTER JOIN Users WITH (nolock) ON Users.DeskCode = ppacreditdesk.TheData And Users.Active = 1 And ppacreditdesk.TheData Is Not NULL And RTRIM(ppacreditdesk.TheData) <> ''"
            .vsql += " WHERE(desk.Branch = 2) And (master.desk In ('10G')) AND (master.closed IS NULL) AND (master.closed IS NULL) OR "
            .vsql += " (desk.Branch = 2) And (master.desk IN ('PPASPECIAL', '104A')) AND (master.closed IS NULL) AND (ppacreditdesk.TheData IS NULL) AND (NOT (ppacreditdesk.TheData IN ('PPASPECIAL', '104A')) OR"
            .vsql += " ppacreditdesk.TheData Is NULL) And ((SELECT tOP (1) entered From payhistory As p3 WITH (nolock) Where (Number = master.number) And (batchtype IN ('PU', 'PC'))"
            .vsql += " Order By entered) BETWEEN @start And @End) Or (desk.Branch = 2) And (master.desk IN ('PPASPECIAL', '104A')) AND (master.closed IS NULL) AND (NOT (ppacreditdesk.TheData IN ('PPASPECIAL', '104A')) OR"
            .vsql += " ppacreditdesk.TheData Is NULL) And (RTRIM(LTRIM(ppacreditdesk.TheData)) = '') AND ((SELECT tOP (1) entered From payhistory As p3 WITH (nolock)"
            .vsql += " Where (Number = master.number) And (batchtype IN ('PU', 'PC')) Order By entered) BETWEEN @start AND @end) OR (desk.Branch = 2) And (master.desk IN ('PPASPECIAL', '104A')) AND (master.closed IS NULL) AND (NOT (ppacreditdesk.TheData IN ('PPASPECIAL', '104A')) OR"
            .vsql += " ppacreditdesk.TheData Is NULL) And (RTRIM(LTRIM(ppacreditdesk.TheData)) Not Like 'A0%') AND ((SELECt TOP (1) entered From payhistory As p3 WITH (nolock)"
            .vsql += " Where (Number = master.number) And (batchtype IN ('PU', 'PC')) Order By entered) BETWEEN @start AND @end) OR (desk.Branch = 2) And (master.desk IN ('PPASPECIAL', '104A')) AND (master.closed IS NULL) AND (NOT (ppacreditdesk.TheData IN ('PPASPECIAL', '104A')) OR"
            .vsql += " ppacreditdesk.TheData Is NULL) And ((SELECT TOP (1) entered From payhistory As p3 WITH (nolock) Where (Number = master.number) And (batchtype IN ('PU', 'PC'))"
            .vsql += " Order By entered) BETWEEN @start AND @end) AND (PPACreditDate.TheData IS NULL) OR (desk.Branch = 2) And (master.desk IN ('PPASPECIAL', '104A')) AND (master.closed IS NULL) AND (NOT (ppacreditdesk.TheData IN ('PPASPECIAL', '104A')) OR"
            .vsql += " ppacreditdesk.TheData Is NULL) And ((SELECT TOP (1) entered From payhistory As p3 WITH (nolock) Where (Number = master.number) And (batchtype IN ('PU', 'PC'))"
            .vsql += " Order By entered) BETWEEN @start AND @end) AND (RTRIM(LTRIM(PPACreditDate.TheData)) = '') OR (desk.Branch = 2) And (master.desk IN ('PPASPECIAL', '104A')) AND (master.closed IS NULL) AND (NOT (ppacreditdesk.TheData IN ('PPASPECIAL', '104A')) OR"
            .vsql += " ppacreditdesk.TheData Is NULL) And ((SELECT TOP (1) entered From payhistory As p3 WITH (nolock) Where (Number = master.number) And (batchtype IN ('PU', 'PC'))"
            .vsql += " Order By entered) BETWEEN @start AND @end) AND (isdate(PPACreditDate.TheData) = 0)ORDER BY master.desk, [File Number]"
        End With

    End Sub


End Class
