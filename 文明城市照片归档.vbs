Option Explicit
'on error resume next 
'******����
Const ImageName2Excel 	= 	"ImageNames2Excel.xls"
Const ExcelTemplate 	= 	"ExcelTemplate.xls"
Const WordTemplate 		= 	"WordTemplate.doc"
Const ExcelGen 			= 	"�����鶽�����.xls"
Const WordGen 			= 	"�����鶽�����.doc"
Const Person1 			= 	"������"
Const Person2 			= 	"�Ⱦ�"
Const PersonFolder1 	= 	"X"
Const PersonFolder2 	= 	"Y"
Const PersonSign1 	= 	"@Hawkinsky"
Const PersonSign2 	= 	"@YoYo"
Const sSheetImageName 	= 	"��Ƭ�����Ӧ"   
Const sSheetPlbName 	= 	"��������������Ƭ"
Const sColPerson		=	"������"		'Column = 1
Const sColImageBaseName		=	"��Ƭ��"		'Column = 2
Const sColImageExt		=	"��Ƭ��չ��"	'Column = 3
Const sColSite			=	"λ��"			'Column = 4
Const sColPlbSite		=	"����λ��"		'Column = 5
Const sColPlbDescription	=	"��������"	'Column = 6


'******VBScript ���õ���Office VBA ����
Const Original_Folder = "ORIGINAL"
Const wdGoToBookmark = 1
Const xlAscending = 1
Const xlDescending = 2
Const xlYes = 1
Const wdCellAlignVerticalCenter = 1
Const wdAutoFitFixed = 2
Const wdSeparateByParagraphs = 0
Const wdCharacter = 1
Const wdExtent = 1
Const wdExtend = 1
Const wdLine = 5
Const wdCell = 12
Const wdTableFormatSimple2 = 2
Const wdAlignParagraphRight = 2
Const wdYellow = 7
Const wdToggle = 9999998
Const wdAlignParagraphCenter = 1
Const wdSentence = 3
Const wdParagraph = 4
Const wdStory = 6
Const wdMove = 0
Const wdNumberGallery = 2
Const wdListApplyToWholeList = 0
Const wdWord10ListBehavior = 2
Const wdWord = 2
Const wdWithInTable = 12
Const wdCollapseStart = 1
'******* ȫ�ֱ���
Dim StdIn	: Set StdIn = WScript.StdIn
Dim StdOut	: Set StdOut = WScript
Dim oShell 	: Set oShell = CreateObject("WScript.Shell")
Dim oFso	: Set oFso = CreateObject("Scripting.FileSystemObject")
'Dim oCFolder: Set oCFolder = oFso.GetFolder(".")
Dim sShellPath	: sShellPath = oShell.CurrentDirectory
Dim sWorkingPath, oWorkingFolder, oFolder, oFile, oFiles
Dim iFileNums, iFileRenamed, iFileRemoved, iImages
Dim sFileName
Dim oExcelApp, oExcelGenWB, oExcelTempWB, oExcelTempSheet
Dim oSheetGen, oSheetImageName, oSheetPlbName
Dim orw, oRange,  oRangePlbName, oRangeGen
Dim aPerson(2,3)
Dim iPerson, sPerson
Dim sY, sM, sD, sDate, sDate_Folder
Dim bFileExists
Dim sBaseName, sExt
aPerson(0,0)	=	Person1
aPerson(0,1)	=	PersonFolder1
aPerson(0,2)	=	PersonSign1
aPerson(1,0)	=	Person2
aPerson(1,1)	=	PersonFolder2
aPerson(1,2)	=	PersonSign2
sDate = Year(Date) & "��" & Month(Date) &"��" & Day(Date) & "��"
sDate_Folder = Year(Date) & Right("0"& Month(Date),2)  & Right("0" & Day(Date),2)
'sWorkingPath = PathCombine(sShellPath, sDate_Folder)
sWorkingPath = sShellPath & "\" & sDate_Folder
Dim strPattern 
strPattern = "IMG_\d{8}_\d{6}\s"						'�ļ���������IMG_20170706_103020 ��ΪAndroid�ֻ���Ƭ
strPattern = strPattern & "|" & "IMG_\d{8}_\d{6}_\d\s"	'�ļ���������IMG_20170706_103020_1 ��ΪAndroid�ֻ���Ƭ
strPattern = strPattern & "|" & "΢��ͼƬ_\d{14}\s"		'�ļ���������΢��ͼƬ_20170706103020 �ַ� ��Ƭ����΢��������ʱ����
strPattern = strPattern & "|" & "_��С��С"				'�ļ��������� _��С��С   ACDSEE ��С��Ƭʱ����
strPattern = strPattern & "|" & "_shrink"				'�ļ��������� _shrink   magick ��С��Ƭʱ�Զ������
strPattern = strPattern & "|" & ".jpg"					'�ļ�������չ��
strPattern = strPattern & "|" & ".tif"					'�ļ�������չ��
strPattern = strPattern & "|" & ".jpeg"					'�ļ�������Ƭ��չ��
strPattern = strPattern & "|" & "\s"					'�ļ����пո�
for iPerson = LBound(aPerson) to UBound(aPerson) -1
	strPattern = strPattern & "|" & aPerson(iPerson,0)
Next

IF not oFso.FolderExists( sWorkingPath ) Then
	CreateWorkingDayFolder
End IF
Set oWorkingFolder = oFso.GetFolder(sWorkingPath)

Dim sInfo 
Dim sStep
sInfo = sInfo & "��������ز��������" & vbCrlf
sInfo = sInfo & "0��ֻ���ɹ���Ŀ¼" & vbCrlf
sInfo = sInfo & "1����ͼƬ����¼��Excel��ȥ��Ȼ���ֹ���д�ص�����������" & vbCrlf
sInfo = sInfo & "2��������д�ĵص�������������������¼����СͼƬ��С������ԭʼͼƬ" & vbCrlf
sInfo = sInfo & "3������ģ������Excel��Word���鱨��" & vbCrlf
sInfo = sInfo & "������ֻ���ɹ���Ŀ¼������Ŀ¼���ɺ󣬰��ֻ����ĵ���Ƭ���͵���������ֹ����浽��Ӧ�Ĺ���Ŀ¼��" & vbCrlf

sStep = InputBox(sInfo, "ѡ�����в���",1)
Select Case sStep
Case "1"
	'��ͼƬ����¼��Excel��ȥ��Ȼ���ֹ���д�ص�����������
	ImagesName2Excel
Case "2"
	'������д�ĵص�������������������¼����СͼƬ��С������ԭʼͼƬ
	RenameImagesFromExcel
Case "3"
	'����ģ������Excel��Word���鱨��
	CreateExcelWordFromTemplate
Case Else
	'������
	'StdOut.Echo "����������ô����������?"
	'��������Ŀ¼
End Select


Function FileIsImage( sExt) 
	sExt = LCase(sExt)
	'StdOut.Echo sExt
	if  "jpg" = sExt or "jpeg" = sExt or "tif" = sExt or "tiff" = sExt or "gif" = sExt or "png" = sExt Then
		FileIsImage = True
	Else
		FileIsImage = False
	End If
End Function

Function RegFileName(ByVal sFileName)
    Dim strReplace : strReplace = ""
    Dim regEx : Set regEx =  new RegExp
	regEx.Global = True
	regEx.MultiLine = False
	regEx.IgnoreCase = False
	regEx.Pattern = strPattern

	If regEx.Test(sFileName) Then
		RegFileName = regEx.Replace(sFileName, strReplace)
	Else
		RegFileName = ""
	End If
End Function

'*****������Ŀ¼·���ϲ�����һ��Ϊ����·�����ڶ���Ϊ���·������\Ϊ��β
Function  PathCombine( sPath1, sPath2)
	If InstrRev( sPath1 , "\") <> len(sPath1) Then	sPath1 = sPath1 + "\"
	If InstrRev( sPath2 , "\") <> len(sPath2) Then	sPath2 = sPath2 + "\"
	PathCombine = sPath1 & sPath2
End Function

'*****��������Ŀ¼
Sub CreateWorkingDayFolder
	Dim bFolderExist
	Dim oShellFolder, aFolder, bFolder , f 
	Set oShellFolder = oFso.GetFolder( sShellPath )
	bFolderExist = oFso.FolderExists( sWorkingPath ) 
	If not bFolderExist Then
		Set aFolder = oShellFolder.SubFolders.Add(sDate_Folder)
		aFolder.SubFolders.Add Original_Folder
		Dim i
		For i = LBound(aPerson) to UBound(aPerson) -1 
			Set bFolder = aFolder.SubFolders.Add(aPerson(i,1))
			bFolder.SubFolders.Add Original_Folder
		Next 
		Set bFolder = nothing
		Set aFolder = nothing
	End If
	Set oShellFolder = nothing
End Sub


	'*****������Ŀ¼��ͼƬ�ļ�����¼����д�뵽 ImageName2Excel ��
	'*****��ʽΪͼƬ�ļ�����������չ���� ��չ�� ��Ƭλ�� ����λ�� �����������ֶ����룩
	'*****��Ƭλ�� ����λ�� ��������  ��Ҫ���ֶ����룩����Ҫ����Ƭ����Ҫ��д
Sub ImagesName2Excel

	'Dim oFile, oFiles
	on   error   resume   next 
	Set oExcelApp = CreateObject("Excel.Application")
    oExcelApp.Visible  =   true
	'sFileName = PathCombine(sWorkingPath , ImageName2Excel)
	'StdOut.Echo sFileName
	bFileExists = oFso.FileExists(sWorkingPath & "\" & ImageName2Excel)
	If bFileExists Then
		Set oExcelTempWB  =  oExcelApp.Workbooks.Open(sWorkingPath & "\" & ImageName2Excel)
		Set oExcelTempSheet = oExcelTempWB.Worksheets( sSheetImageName )
	Else
		Set oExcelTempWB  =  oExcelApp.Workbooks.Add
		Set oExcelTempSheet = oExcelTempWB.Worksheets(2)
		With oExcelTempSheet
				'Const sSheetImageName 	= 	"��Ƭ�����Ӧ"   
				'Const sSheetPlbName 	= 	"��������������Ƭ"
				'Const sColPerson		=	"������"		'Column = 1
				'Const sColImageBaseName		=	"��Ƭ��"		'Column = 2
				'Const sColImageExt		=	"��Ƭ��չ��"	'Column = 3
				'Const sColSite			=	"λ��"			'Column = 4
				'Const sColPlbSite		=	"����λ��"		'Column = 5
				'Const sColPlbDescription	=	"��������"	'Column = 6
			.Name = sSheetPlbName
			.Cells(1,1).value = sColPerson 
			.Cells(1,2).value = sColImageBaseName 
			.Cells(1,3).value = sColImageExt
			.Cells(1,4).value = sColSite		
			.Cells(1,5).value = sColPlbSite
			.Cells(1,6).value = sColPlbDescription		
			.UsedRange.EntireColumn.Autofit()
		End With
		Set oExcelTempSheet = oExcelTempWB.Worksheets( 1 )
		With oExcelTempSheet
			.Activate
			.Name = sSheetImageName
			.Cells(1,1).value = sColPerson 
			.Cells(1,2).value = sColImageBaseName 
			.Cells(1,3).value = sColImageExt 
			.Cells(1,4).value = sColSite 
			.Cells(1,5).value = sColPlbSite 
			.Cells(1,6).value = sColPlbDescription 
		End With
	End If
	oExcelTempSheet.Activate

	iImages = 0
	'StdOut.Echo LBound(aPerson) & "; " & UBound(aPerson) & "   " & oWorkingFolder.Name
	For Each oFolder in oWorkingFolder.SubFolders
		'StdOut.Echo oFolder.Name
		iPerson = FoundPersonByFolder( oFolder.Name ) 
		if iPerson <> -1 Then 
			Set oFiles = oFolder.Files
			For Each oFile in oFiles
				sFileName = oFile.Name
				sBaseName = Left(sFileName, InstrRev(sFileName,".") - 1)
				sExt = Right(sFileName,len(sFileName) - InstrRev(sFileName,"."))
				'StdOut.Echo "filename=" & sFileName & "  basename=" & sBaseName & "  basename=" & sExt 
				If FileIsImage(sExt) Then 
					oExcelTempSheet.Cells(2+iImages, 1).value =  aPerson(iPerson,0)
					oExcelTempSheet.Cells(2+iImages, 2).value =  sBaseName 
					oExcelTempSheet.Cells(2+iImages, 3).value =  sExt 
					iImages = iImages + 1
				End If 
			Next
		End If 
	Next
	oExcelTempSheet.UsedRange.EntireColumn.Autofit()
	
	If bFileExists Then
		oExcelTempWB.Save
	Else
		oExcelTempWB.SaveAs sWorkingPath & "\" & ImageName2Excel
	End If

	oExcelApp.Application.quit
	Set oExcelTempSheet = nothing
	set oExcelTempWB  =   nothing
	Set oExcelApp  =   nothing
End Sub

'����Ŀ¼�������ң�����ֵΪ��������ڼ���������ֵΪ-1��ʾ�Ҳ���
Function FoundPersonByFolder( sFolerName )
	Dim i
	FoundPersonByFolder = -1
	For i = LBound(aPerson) to UBound(aPerson) -1 
		if aPerson(i,1) = sFolerName Then Exit For
	Next 
	FoundPersonByFolder = i
End Function

'�������������ң�����ֵΪ��������ڼ���������ֵΪ-1��ʾ�Ҳ���
Function FoundPerson( sPerson )
	Dim i
	FoundPerson = -1
	For i = LBound(aPerson) to UBound(aPerson) -1 
		if aPerson(i,0) = sPerson Then Exit For
	Next 
	FoundPerson = i
End Function

Sub RenameImagesFromExcel
	Dim img
	Dim msgs 

	Set oExcelApp = CreateObject("Excel.Application")
	oExcelApp.Visible  =   True
	Set oExcelTempWB  =  oExcelApp.Workbooks.Open(sWorkingPath & "\" & ImageName2Excel)
	Set oSheetImageName = oExcelTempWB.Worksheets( sSheetImageName)
	Set oSheetPlbName = oExcelTempWB.Worksheets( sSheetPlbName )
	oSheetPlbName.Activate
	Dim iRowNums
	iRowNums = oSheetImageName.UsedRange.Rows.Count
	'StdOut.Echo iRowNums
	Set oRange = oSheetImageName.UsedRange
	Set oRangePlbName = oSheetPlbName.UsedRange
	Set img = CreateObject("ImageMagickObject.MagickImage.1")
				'Const sSheetImageName 	= 	"��Ƭ�����Ӧ"   
				'Const sSheetPlbName 	= 	"��������������Ƭ"
				'Const sColPerson		=	"������"		'Column = 1
				'Const sColImageBaseName		=	"��Ƭ��"		'Column = 2
				'Const sColImageExt		=	"��Ƭ��չ��"	'Column = 3
				'Const sColSite			=	"λ��"			'Column = 4
				'Const sColPlbSite		=	"����λ��"		'Column = 5
				'Const sColPlbDescription	=	"��������"	'Column = 6
	Dim sNewBaseName, sFullPath, sFullNewBaseName, sCurrentDirectory, Command
	'sCurrentDirectory = oShell.CurrentDirectory
	'oShell.CurrentDirectory = sWorkingPath
	For Each orw In oRange.Rows
		If orw.Row <> 1 then
			sPerson = oRange.Cells( orw.Row , 1)
			iPerson = FoundPerson( sPerson) 
			If iPerson <> -1 then '�ҵ�
				sBaseName = oRange.Cells( orw.Row , 2)
				sExt = oRange.Cells( orw.Row , 3)
				sNewBaseName = oRange.Cells( orw.Row , 4) & " " & oRange.Cells( orw.Row , 5) & oRange.Cells( orw.Row , 6)
				sFullPath = PathCombine(sWorkingPath ,aPerson(iPerson,1))
				'StdOut.Echo aPerson(iPerson,1) & ":" & sFullPath
				sFullNewBaseName = sPerson &" " & sBaseName & " " & sNewBaseName 
				'StdOut.Echo  sFullPath & sBaseName  & "." & sExt
				If oFso.FileExists(sFullPath & sBaseName  & "." & sExt) Then
					Set oFile = oFso.GetFile(sFullPath & sBaseName & "." & sExt)
					'StdOut.Echo  oFile.Path & "  filename:" & oFile.Name
					If trim(sNewBaseName) <> "" Then  
						'���������ݿ��� ��ӵ� WorkSheet(sSheetPlbName)��
						oSheetPlbName.Cells(oRangePlbName.Rows.Count+1+iFileRenamed,1).value = oRange.Cells( orw.Row,1).value
						oSheetPlbName.Cells(oRangePlbName.Rows.Count+1+iFileRenamed,2).value = oRange.Cells( orw.Row,2).value
						oSheetPlbName.Cells(oRangePlbName.Rows.Count+1+iFileRenamed,3).value = oRange.Cells( orw.Row,3).value
						oSheetPlbName.Cells(oRangePlbName.Rows.Count+1+iFileRenamed,4).value = oRange.Cells( orw.Row,4).value
						oSheetPlbName.Cells(oRangePlbName.Rows.Count+1+iFileRenamed,5).value = oRange.Cells( orw.Row,5).value
						oSheetPlbName.Cells(oRangePlbName.Rows.Count+1+iFileRenamed,6).value = oRange.Cells( orw.Row,6).value

						'�����Ƭ�����������򱣴浽����Ŀ¼��ɾ����ƬEXIF��Ϣ��ͬʱ����ΪsPerson & " " & sBaseName & " " & sNewBaseName & "." & sExt 
						'*****��СͼƬ
						msgs = img.Convert( oFile.Path, _
							"-strip","-resize","226x170", _
							"-fill", "white", "-pointsize", "10", _
							"-gravity","southeast","-annotate", "+0+0", aPerson(iPerson,2), _
							sWorkingPath & sFullNewBaseName & "_shrink." & sExt)
						oFile.Copy PathCombine(sWorkingPath ,Original_Folder) & sFullNewBaseName & "." & sExt
						oFile.Move PathCombine(sFullPath ,Original_Folder)
						iFileRenamed = iFileRenamed + 1
					Else
						'�������Ƭ˵�����ƶ���Original_Folder 
						'StdOut.Echo sCFolder & "\" & Original_Folder & "\"
						oFile.Move PathCombine(PathCombine(sWorkingPath ,aPerson(iPerson,1)) ,Original_Folder)
						iFileRemoved = iFileRemoved + 1
					End If
					Set oFile = nothing
				End If
			End If
		End If
	Next 
	
	oSheetPlbName.UsedRange.EntireColumn.Autofit()
	oExcelTempWB.Save
	oExcelTempWB.Close
	oExcelApp.Application.quit
	Set img = nothing
	Set oSheetImageName = nothing
	set oExcelTempWB  =   nothing
	Set oExcelApp  =   nothing
End Sub

'����ģ������Excel��Word���鱨��
Sub CreateExcelWordFromTemplate
	Dim Header , OrderCustom, MatchCase, Orientation, ColumnC
	Dim plbSite, plbDescription, lastplbSite
	Dim oWordApp, oWordAppBookmak, oWordAppRange, oWordGenDoc, oSelection
	Dim sBookmarkName, Bookmark
	Dim strPath
	Dim strName
	Dim strFolderPath
	Dim oCell
	Dim iImageNums, i, j, iRowNums, iRow, s1, s2, oShape

	bFileExists = oFso.FileExists(oShell.CurrentDirectory & "\" & ExcelTemplate)
	If not bFileExists Then 
		StdOut.Echo "Excelģ�����ģ�"
		Exit Sub
	Else
		Set oFile = oFso.GetFile(oShell.CurrentDirectory & "\" & ExcelTemplate)
	End If
	bFileExists = oFso.FileExists(sWorkingPath & "\" & sDate & ExcelGen)
	If not bFileExists Then 
		oFile.Copy sWorkingPath  & "\" & sDate & ExcelGen, True
	End If
	
	Set oExcelApp = CreateObject("Excel.Application")
	oExcelApp.Visible  =   true
	Set oExcelTempWB  =  oExcelApp.Workbooks.Open(sWorkingPath & "\" & IMAGENAME2EXCEL)
	Set oExcelGenWB  =  oExcelApp.Workbooks.Open( sWorkingPath & "\" & sDate & ExcelGen)
	Set oSheetPlbName = oExcelTempWB.Worksheets( sSheetPlbName )
	oSheetPlbName.Activate
	Set oSheetGen = oExcelGenWB.Worksheets( "Sheet1" )
	oSheetGen.Activate
	oSheetGen.Cells(2,4).value = sDate
	sPerson = ""
	for iPerson = LBound(aPerson) to UBound(aPerson) -1
		If sPerson = "" Then 
			sPerson = aPerson(iPerson,0)
		Else
			sPerson = sPerson & "��" & aPerson(iPerson,0)
		End If 
	Next
	oSheetGen.Cells(2,2).value = sPerson     '��Ա��
	'Const sSheetImageName 	= 	"��Ƭ�����Ӧ"   
	'Const sSheetPlbName 	= 	"��������������Ƭ"
	'Const sColPerson		=	"������"		'Column = 1
	'Const sColImageBaseName		=	"��Ƭ��"		'Column = 2
	'Const sColImageExt		=	"��Ƭ��չ��"	'Column = 3
	'Const sColSite			=	"λ��"			'Column = 4
	'Const sColPlbSite		=	"����λ��"		'Column = 5
	'Const sColPlbDescription	=	"��������"	'Column = 6
	For iRow = 2 to oSheetPlbName.UsedRange.Rows.Count
		s1 = oSheetPlbName.UsedRange.Cells( iRow , 4)
		s2 = oSheetPlbName.UsedRange.Cells( iRow , 5) & oSheetPlbName.UsedRange.Cells( iRow , 6)
		oSheetGen.Cells(iRow + 2,1).value = iRow - 1
		oSheetGen.Cells(iRow + 2,2).value = "25"
		oSheetGen.Cells(iRow + 2,3).value = s1
		oSheetGen.Cells(iRow + 2,4).value = s2
	Next 
	'StdOut.Echo "��������������Ƭ��" & oSheetGen.UsedRange.Rows.Count & "��"
	'oSheetGen.UsedRange.EntireColumn.Autofit()
	oExcelGenWB.Save
	oExcelTempWB.Close '�ر�IMAGENAME2EXCEL
	Set oExcelTempWB = nothing 
	
	'Sort ����
	Header = 1 'use first row as column headings - (default) 1 = Yes, (?)2 = No
	OrderCustom = 1 'index of custom sort order from Sort Options dialog box - (default) 1 = Normal
	MatchCase = False 'True = case sensitive, (default) False = ignore case
	Orientation = 1 '(default) 1 = top to bottom, (?)2 = left to right
	Set ColumnC = oSheetGen.Range("C3")
	Set oRangeGen = oSheetGen.UsedRange
	'StdOut.Echo "ʹ��������" & oSheetGen.UsedRange.Rows.Count
	Set oRange = oRangeGen.Range("C3:D"& oRangeGen.Rows.Count)
	'oRange.Select
	oRange.Sort ColumnC, xlAscending, , , , , , Header, OrderCustom,  MatchCase, Orientation
	Set ColumnC = nothing
	Set oRange = nothing
	Set oRangeGen = nothing
	oExcelGenWB.Save

	'***** ����Word
	'***** ��ʽΪͼƬ�ļ�����������չ����  ��չ��  ��Ƭ˵�����ֶ����룩
	Set oRangeGen = oSheetGen.UsedRange
	'StdOut.Echo "oRangeGenʹ��������" & oRangeGen.Rows.Count
	plbSite = ""
	lastplbSite = ""
	bFileExists = oFso.FileExists(oShell.CurrentDirectory  & "\" & WordTemplate)
	If not bFileExists Then 
		StdOut.Echo "Wordģ�����ģ�"
		Exit Sub
	Else
		Set oFile = oFso.GetFile(oShell.CurrentDirectory  & "\" & WordTemplate)
	End If
	bFileExists = oFso.FileExists(sWorkingPath  & "\" & sDate & WordGen)
	If not bFileExists Then 
		oFile.Copy sWorkingPath  & "\" & sDate & WordGen, True
	End If

	Set oWordApp = CreateObject("Word.Application")
	oWordApp.Visible = True
	Set oWordGenDoc = oWordApp.Documents.Open( sWorkingPath  & "\" & sDate & WordGen)
	oWordGenDoc.Activate
	Set oSelection = oWordApp.Selection

	For Each Bookmark in oWordGenDoc.Bookmarks
		Select Case Bookmark.Name
		Case "����"
			'Bookmark.Range.Text = sDate
			'.Goto wdGoToBookmark, ,  , Bookmark.Name
			Bookmark.Select
			Set oSelection = oWordApp.Selection
			oSelection.TypeText sDate
		Case "·��"
			'***** Excel ����C���в��ظ���ֵ
			Set oRange = oRangeGen.Range("C4:C"& oRangeGen.Rows.Count)
			'oSelection.Goto wdGoToBookmark, ,  , Bookmark.Name
			Bookmark.Select
			Set oSelection = oWordApp.Selection
			iRowNums = oRange.Rows.Count
			plbSite = ""
			lastplbSite = ""
			For Each orw in oRange.Rows
				plbSite = oRangeGen.Cells(orw.Row,3).value
				If plbSite <> lastplbSite Then 
					lastplbSite = plbSite
					If orw.Row <> oRangeGen.Rows.Count Then 
						oSelection.TypeText plbSite  & "��"
					Else 
						oSelection.TypeText plbSite  & "��"
					End If
				End If
			Next
			'sRoute = Left(sRoute, len(sRoute)-1) & "��" 
			'Bookmark.Range.Text = sRoute 
			oSelection.TypeBackspace
			oSelection.TypeText "��"
			Set oRange = nothing
		Case "��Ա" 
			sPerson = ""
			for iPerson = LBound(aPerson) to UBound(aPerson) -1
				If sPerson = "" Then 
					sPerson = aPerson(iPerson,0)
				Else
					sPerson = sPerson & "��" & aPerson(iPerson,0)
				End If 
			Next
			Bookmark.Select
			Set oSelection = oWordApp.Selection
			oSelection.TypeText sPerson
		Case "����" 
			'oSelection.Goto wdGoToBookmark, ,  , Bookmark.Name
			Bookmark.Select
			Set oSelection = oWordApp.Selection
			'***** Excel ����C���в��ظ���ֵ
			Set oRange = oRangeGen.Range("C4:D"& oRangeGen.Rows.Count)

			plbDescription = ""

			iRowNums = oRange.Rows.Count 
			'StdOut.Echo "����iRowNums = " & iRowNums 
			iRow = 0
			For Each orw in oRange.Rows
				'StdOut.Echo "����orw.Row = " & orw.Row 
				plbSite = oRangeGen.Cells(orw.Row,3).value
				plbDescription = oRangeGen.Cells(orw.Row,4).value
				If plbSite <> lastplbSite Then 
					iRow = iRow + 1
					lastplbSite = plbSite
					'�����һ���򲻻�һ�У���������һ��
					If orw.Row <> 4 Then 
						oSelection.TypeBackspace ' ɾ��������
						oSelection.TypeText "��"
						oSelection.TypeParagraph
					End If
					oSelection.TypeText plbSite & "��"
					If orw.Row <> oRangeGen.Rows.Count Then 
						oSelection.TypeText plbDescription & "��"
					Else 
						oSelection.TypeText plbDescription & "��"
					End If
				Else 
					If orw.Row <> oRangeGen.Rows.Count Then 
						oSelection.TypeText plbDescription & "��"
					Else 
						oSelection.TypeText plbDescription & "��"
					End If
				End If
			Next
			'****�����Եص���б��
			Bookmark.Select
			oSelection.MoveDown wdParagraph, iRow, wdExtend
			oSelection.Range.ListFormat.ApplyListTemplate oWordApp.ListGalleries( _
				wdNumberGallery).ListTemplates(2), False, _
				wdListApplyToWholeList, wdWord10ListBehavior

			Set oRange = nothing
		Case "ͼƬ" 
			'*****����ͼƬ ��ת���ɱ�� 
			'oSelection.Goto wdGoToBookmark, ,  , Bookmark.Name
			Bookmark.Select
			Set oSelection = oWordApp.Selection
			oWordApp.ScreenUpdating = False
			'loops through each file in the directory and prints their names and path
			iImageNums = 0
			For Each oFile In oWorkingFolder.Files
				'get file path
				strPath = oFile.Path
				strName = oFile.Name
				sExt = oFso.GetExtensionName(oFile)
				If FileIsImage(sExt) Then
					'insert the image
					oSelection.InlineShapes.AddPicture strPath, False, True
					oSelection.TypeText RegFileName(strName)
					oSelection.TypeParagraph
					iImageNums = iImageNums + 1
				End If
			Next 
			oWordApp.ScreenRefresh
			'oSelection.Goto wdGoToBookmark,  ,  , Bookmark.Name
			'StdOut.Echo iImageNums
			Bookmark.Select
			oSelection.MoveDown wdParagraph, iImageNums, wdExtend
			oSelection.Range.ConvertToTable wdSeparateByParagraphs,ABS((iImageNums+1 )/2), 2, oWordApp.CentimetersToPoints(8.0)
			
			oSelection.Tables(1).Select
			oSelection.Paragraphs.Alignment = wdAlignParagraphCenter
			oSelection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
			Dim oShapes
			Set oShapes = oWordGenDoc.InlineShapes
			For Each oShape in oShapes
				oShape.Select
				oSelection.Collapse wdCollapseStart
				If oShape.Height > oShape.Width and oSelection.Information(wdWithInTable) Then
                    oSelection.MoveRight wdWord, 1
                    oSelection.TypeParagraph
				End If
			Next
			oWordApp.ScreenUpdating = True
			Set oShapes = nothing
		Case else
			StdOut.Echo "�в���ʶ����ǩ����Ҫ�����ˡ�" & Bookmark.Name
			'Exit Sub
		End select
	Next
	
	oExcelGenWB.Close
	
	oWordGenDoc.Save 
	oWordGenDoc.Close
	oWordApp.Application.Quit
	oExcelApp.Application.Quit
	Set oSheetPlbName = nothing
	Set oSheetGen = nothing
	Set oExcelTempWB = nothing
	Set oExcelGenWB  =   nothing
	Set oWordGenDoc = nothing
	Set oWordApp = nothing
	Set oExcelApp  =   nothing
End Sub


Set oWorkingFolder = nothing
Set oFso = nothing
WScript.Quit(0)