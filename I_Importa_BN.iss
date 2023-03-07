'Desarrollado por Jorge M. Chávez
'Fecha: 01/03/2023

Sub Main
	IgnoreWarning(True)
	Call ReportReaderImport()	'D:\RUC1\DATA\Archivos fuente.ILB\2022_BN.pdf
	Call AppendField()	'I_BN2022.IMD
	Call Summarization()	'I_BN2022.IMD
	Client.CloseAll
	Dim pm As Object
	Dim SourcePath As String
	Dim DestinationPath As String
	Set SourcePath = Client.WorkingDirectory
	Set DestinationPath = "D:\RUC1\DATA\_EECC"
	Client.RunAtServer False
	Set pm = Client.ProjectManagement
	pm.MoveDatabase SourcePath + "I_BN2022.IMD", DestinationPath
	pm.MoveDatabase SourcePath + "I.1_Resumen_BN.IMD", DestinationPath
	Set pm = Nothing
	Client.RefreshFileExplorer
End Sub


' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport
	dbName = "I_BN2022.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\BN_CTA_CTE.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\2022_BN.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Anexar campo
Function AppendField
	Set db = Client.OpenDatabase("I_BN2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "PERIODO"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@RIGHT(FECHA;4)+@MID(FECHA;4;2)"
	field.Length = 6
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Análisis: Resumen
Function Summarization
	Set db = Client.OpenDatabase("I_BN2022.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "PERIODO"
	task.AddFieldToSummarize "CUENTA"
	task.AddFieldToTotal "CARGOS"
	task.AddFieldToTotal "ABONOS"
	dbName = "I.1_Resumen_BN.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function