Attribute VB_Name = "doc�����޸�"
Dim hege As Variant

Sub AgetHege()

    '''�������'''''''''''''''''''''''''''''''''''''''''''
    Sheets("�ڵغϸ����").Select
    Range("D:F").ClearContents
    
    mainDir = ThisWorkbook.Path
    xlsPath = mainDir + "\" + "�ڵغϸ������.xls"
    
    '''����ʱ��'''''''''''''''''''''''''''''''''''''''''''
    surveyTime = InputBox("���������ʱ�䣬" + Chr(13) + "���磺2020��08��07��", "�޸ĵ���ʱ��", "2020��08��09��")
    
    '''�ϴ��б�'''''''''''''''''''''''''''''''''''''''''''
    targetHukou = InputBox("������������������кϴ壬���á��ָ���" + Chr(13) + "��1������ʡ��غ�������޴�" + Chr(13) + "��2������ʡ��غ�������ľ���塢�������", "�����ʸ��ж�", "����ʡ��ع������Ҹ���")
    Dim targetHukouList()
    Dim cunzhenIndex()
    k = 0
    If InStr(targetHukou, "��") Then
        targetHukou = Replace(targetHukou, "��", "")
        zhenIndex = InStr(targetHukou, "��")
        ReDim Preserve cunzhenIndex(k)
        cunzhenIndex(k) = zhenIndex + 1
        k = k + 1
        zhenName = Left(targetHukou, zhenIndex)
        
        cunloc = 0
        Do
            cunloc = InStr(cunloc + 1, targetHukou, "��")
            If cunloc > 0 Then
                ReDim Preserve cunzhenIndex(k)
                cunzhenIndex(k) = cunloc + 1
                k = k + 1
            End If
        Loop Until cunloc = 0
        
        For j = 0 To UBound(cunzhenIndex) - 1
            ReDim Preserve targetHukouList(j)
            targetHukouList(j) = zhenName + Mid(targetHukou, cunzhenIndex(j), cunzhenIndex(j + 1) - cunzhenIndex(j))
        Next
    Else
        ReDim targetHukouList(0)
        targetHukouList(0) = targetHukou
    End If
    
    With Sheets("�ڵغϸ����")
        
        Dim wordApp As Word.Application
        Dim wordDoc As Word.Document
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = True
        wordApp.Activate
        
        lastRow = .Range("A1").End(xlDown).Row
        For i = 2 To lastRow
            dirName = .Cells(i, 1) + .Cells(i, 2)
            'doc03
            If Dir(mainDir + "\" + dirName + "\03*.doc") <> "" Then
                doc03 = mainDir + "\" + dirName + "\" + Dir(mainDir + "\" + dirName + "\03*.doc")
                Set wordDoc = wordApp.Documents.Open(doc03)
                wordDoc.Activate
                Call BdealDoc03(wordDoc, mainDir, surveyTime, targetHukouList, .Cells(i, 3))
                wordDoc.Close True
                Set wordDoc = Nothing
            Else
                .Cells(1, 4) = "�Ҳ���03"
                .Cells(i, 4) = "��"
            End If
            'doc07
            If Dir(mainDir + "\" + dirName + "\07*.doc") <> "" Then
                doc07 = mainDir + "\" + dirName + "\" + Dir(mainDir + "\" + dirName + "\07*.doc")
                Set wordDoc = wordApp.Documents.Open(doc07)
                wordDoc.Activate
                Call BdealDoc07(wordDoc, mainDir, targetHukouList)
                wordDoc.Close True
                Set wordDoc = Nothing
            Else
                .Cells(1, 5) = "�Ҳ���07"
                .Cells(i, 5) = "��"
            End If
            'doc02
            Dim doc02s()
            j = 0
            If Dir(mainDir + "\" + dirName + "\02*.doc") <> "" Then
                doc02 = Dir(mainDir + "\" + dirName + "\02*.doc")
                Do While doc02 <> ""
                    ReDim Preserve doc02s(j)
                    doc02s(j) = doc02
                    j = j + 1
                    doc02 = Dir()
                Loop
                For Each doc0 In doc02s
                    Set wordDoc = wordApp.Documents.Open(mainDir + "\" + dirName + "\" + doc0)
                    wordDoc.Activate
                    Call BdealDoc02(wordDoc, mainDir, targetHukouList, UBound(doc02s), .Cells(i, 3))
                    wordDoc.Close True
                    Set wordDoc = Nothing
                Next
            Else
                Cells(1, 6) = "�Ҳ���02"
                Cells(i, 6) = "��"
            End If
        Next
    End With
    
    wordApp.Quit
    Set wordApp = Nothing
    
    ThisWorkbook.SaveAs mainDir + "\" + "�doc�����޸Ľ��.xls"
    
End Sub

Sub BdealDoc03(ByVal wordDoc As Word.Document, mainDir, surveyTime, targetHukouList, hege)
    
    Dim wordApp As Word.Application
    Set wordApp = GetObject(, "Word.Application")
    Set wordApp = wordDoc.Parent
    
'    CurrentDocStart = wordApp.Selection.GoTo(what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1).Start
'    CurrentDocEnd = wordApp.Selection.GoTo(what:=wdGoToLine, Which:=wdGoToLast).End
    
    ZPath = mainDir + "\" + Mid(targetHukouList(0), 10, 4)

    '''PPP1''''''''''''''''''''''''''''''''''''''''
    wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    CurrentPageStart = wordApp.Selection.Start
    CurrentPageEnd = wordApp.Selection.Goto(what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=2).Start
    Set myRange = wordDoc.Range(CurrentPageStart, CurrentPageEnd)
    If myRange.ShapeRange.Count <> 0 Then
        myRange.ShapeRange.Select
        myRange.ShapeRange.Delete
    End If
    
    wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
    '''���鵥λ
    wordDoc.Paragraphs(19).Range.Select
    With wordApp.Selection
        .MoveRight Unit:=wdCell, Count:=1
        .TypeBackspace
        .TypeText Text:="�����Ȼ��Դ��" + Chr(13) + "����ʡ��ز�湤�����޹�˾"
    End With
    
    '''�������
    NFZAngle = Int(Rnd * 135 + 1)
    NFZTop = Int(Rnd * 25 + 385)
    NFZLeft = Int(Rnd * 80 + 160)
    Set NFZ = wordDoc.Shapes.AddPicture(Filename:=ZPath + "\ũ����.png", LinkToFile:=False, SaveWithDocument:=True) '����ͼƬ
    With NFZ
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .IncrementTop NFZTop
        .IncrementLeft NFZLeft
        .Rotation = NFZAngle
        .WrapFormat.Type = wdWrapBehind
        .ZOrder 5
        .Select
        .Name = "pRed 1"
    End With
    
    '''����ʱ��
    nian = Left(surveyTime, 4)
    yue = Mid(surveyTime, 6, 2)
    ri = Mid(surveyTime, 9, 2)
    
    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "����ʱ�䣺"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute
    wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=2
    wordApp.Selection.Words(1).Text = nian
    wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=5
    wordApp.Selection.Words(1).Text = yue
    wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=3
    wordApp.Selection.Words(1).Text = ri
    
    '''�ж��Ƿ�໧��'''''''''''''''''''''''''''''''''''''
    tableCount = wordDoc.Tables.Count
    If tableCount > 6 Then
        QLRNum = tableCount - 5
    Else
        QLRNum = 1
    End If
    
    For i = 2 To QLRNum + 1
        '''PPP2'''''''''''''''''''''''''''''''''''''''''''
        wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=i
        CurrentPageStart = wordApp.Selection.Start
        CurrentPageEnd = wordApp.Selection.Goto(what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=i + 1).Start
        Set myRange = wordDoc.Range(CurrentPageStart, CurrentPageEnd)
        If myRange.ShapeRange.Count <> 0 Then
            myRange.ShapeRange.Select
            myRange.ShapeRange.Delete
        End If
        
        wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=i
        TableName2 = wordDoc.Tables(i).Cell(1, 1).Range.Text
        If Left(TableName2, 7) = "��ͥ��Ա�����" Then
            Do While wordApp.Selection.Text = Chr(13) Or wordApp.Selection.Text = " "
                wordApp.Selection.Delete
            Loop
            Set p2 = wordDoc.Shapes.AddPicture(Filename:=ZPath + "\p2.png", LinkToFile:=False, SaveWithDocument:=True, Width:=CentimetersToPoints(21), Height:=CentimetersToPoints(29.7), Anchor:=wordApp.Selection.Range) '����ͼƬ
            With p2
                .WrapFormat.Type = wdWrapBehind
                .ZOrder 5
                .Name = "p2"
                .Select
            End With
            wordApp.Selection.ShapeRange.Align msoAlignLefts, True
            wordApp.Selection.ShapeRange.Align msoAlignTops, True
            wordApp.Selection.ShapeRange.Align msoAlignCenters, True

            wordApp.Selection.Find.ClearFormatting
            With wordApp.Selection.Find
                .Text = "1���û������й��м�ͥ��Ա"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
            End With
            wordApp.Selection.Find.Execute
            BZ1 = wordApp.Selection.Cells(1).RowIndex
            
            Set range2 = wordDoc.Range( _
                wordDoc.Tables(i).Cell(2, 2).Range.Start, _
                wordDoc.Tables(i).Cell(BZ1 - 1, 2).Range.End)
            range2.Cells.Height = CentimetersToPoints(1.12)
            
            wordApp.Selection.Find.ClearFormatting
            With wordApp.Selection.Find
                .Text = "��  ��  ��"
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
            End With
            wordApp.Selection.Find.Execute
            wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=1
            
            If QLRNum > 1 Then
                wordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 7
            Else
                wordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 6
            End If
            
            '''��ע����'''''''''''''''''''''''''''''''''''''''''''
            lastR = wordDoc.Tables(i).Rows.Count - 4
            nameHuzhu = wordDoc.Tables(i).Cell(2, 3).Range.Text
            nameHuzhu = Replace(nameHuzhu, Chr(13) + "", "")
            numRen = 0
            numHuzhu = 0
            For r2 = 7 To lastR
                wordDoc.Tables(i).Cell(r2, 5).Range.Select
                nameRela = wordDoc.Tables(i).Cell(r2, 5).Range.Text
                nameRela = Replace(nameRela, Chr(13) + "", "")
                If nameRela <> "" Then
                    numRen = numRen + 1
                    If nameRela = "����" Or nameRela = "����" Then
                        numHuzhu = numHuzhu + 1
                    End If
                    wordDoc.Tables(i).Cell(r2, 7).Range.Text = "����" + nameHuzhu
                Else
                    Exit For
                End If
            Next r2
            
            '''�����ʸ��ж�'''''''''''''''''''''''''''''''''''''''''''

            ''''''��ǰ�������ڵ�''''''
            hukouSuozaidi = wordDoc.Tables(i).Cell(4, 3).Range.Text
            cunIndex = InStr(hukouSuozaidi, "��")
            currentHukou = Left(hukouSuozaidi, cunIndex)
            
            Dim benCunzu As Variant
            For j = 0 To UBound(targetHukouList)
                If currentHukou = targetHukouList(j) Then
                    benCunzu = 1 '�Ǳ�����
                    Exit For
                Else
                    benCunzu = 0 '�Ǳ�����
                End If
            Next
            If benCunzu = 0 Then '�Ǳ�����
                '''լ�����ʸ�Ȩ2
                wordDoc.Tables(i).Select
                wordApp.Selection.Find.ClearFormatting
                With wordApp.Selection.Find
                    .Text = "��լ�����ʸ�Ȩ"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                End With
                wordApp.Selection.Find.Execute
                wordApp.Selection.MoveLeft Unit:=wdCharacter, Count:=2
                wordApp.Selection.TypeBackspace
                wordApp.Selection.TypeText Text:="0"
                
                wordApp.Selection.Find.ClearFormatting
                With wordApp.Selection.Find
                    .Text = "�˷���լ���طֻ���������"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                End With
                wordApp.Selection.Find.Execute
                wordApp.Selection.MoveLeft Unit:=wdCharacter, Count:=2
                wordApp.Selection.TypeBackspace
                wordApp.Selection.TypeText Text:="0"
                
                wordApp.Selection.Find.ClearFormatting
                With wordApp.Selection.Find
                    .Text = "3�������ʸ��жϣ�"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                End With
                wordApp.Selection.Find.Execute
                
                wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=2
                wordApp.Selection.TypeBackspace
                wordApp.Selection.InsertSymbol Font:="����", CharacterNumber:=9633, Unicode:=True
                wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=13
                wordApp.Selection.TypeBackspace
                wordApp.Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-4014, Unicode:=True

            
                '''լ�����ʸ�Ȩ7
                wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=i + 5
                wordApp.Selection.Find.ClearFormatting
                With wordApp.Selection.Find
                    .Text = "��լ�����ʸ�Ȩ"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                End With
                wordApp.Selection.Find.Execute
                wordApp.Selection.MoveLeft Unit:=wdCharacter, Count:=2
                wordApp.Selection.TypeBackspace
                wordApp.Selection.TypeText Text:="0"
            End If
            
            chrBZ2 = wordDoc.Tables(i).Cell(BZ1 + 1, 2).Range.Characters.Count
            If chrBZ2 > 68 Then
                p2.IncrementTop 20
            End If
        End If
    Next
    
    '''PPP3'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=QLRNum + 2
    CurrentPageStart = wordApp.Selection.Start
    CurrentPageEnd = wordApp.Selection.Goto(what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=QLRNum + 3).Start
    Set myRange = wordDoc.Range(CurrentPageStart, CurrentPageEnd)
    If myRange.ShapeRange.Count <> 0 Then
        myRange.ShapeRange.Select
        myRange.ShapeRange.Delete
    End If
    
    wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=QLRNum + 2
    Do While wordApp.Selection.Text = Chr(13) Or wordApp.Selection.Text = " "
        wordApp.Selection.Delete
    Loop
    '''������'''''''''''''''''''''''''''''''''''''''''''''''''
    Set p3 = wordDoc.Shapes.AddPicture(Filename:=ZPath + "\p3.png", LinkToFile:=False, SaveWithDocument:=True, Width:=CentimetersToPoints(21), Height:=CentimetersToPoints(29.7), Anchor:=wordApp.Selection.Range) '����ͼƬ
    With p3
        .WrapFormat.Type = wdWrapBehind
        .ZOrder 5
        .Name = "p3"
        .Select
    End With
    wordApp.Selection.ShapeRange.Align msoAlignLefts, True
    wordApp.Selection.ShapeRange.Align msoAlignTops, True
    wordApp.Selection.ShapeRange.Align msoAlignCenters, True
    
    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "��׼�ļ�"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute
    wordApp.Selection.Cells.Height = CentimetersToPoints(1.8)
    
    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "Ԥ���"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute
    YBH = wordApp.Selection.Cells(1).RowIndex
    
    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "������Դ"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute
    FWLY = wordApp.Selection.Cells(1).RowIndex
    
    '''����-����
    Set range31 = wordDoc.Range( _
        wordDoc.Tables(QLRNum + 2).Cell(FWLY - 5, 2).Range.Start, _
        wordDoc.Tables(QLRNum + 2).Cell(FWLY - 1, 2).Range.End)
    range31.Cells.Height = CentimetersToPoints(0.6)
    
    '''����
    Set range32 = wordDoc.Range( _
        wordDoc.Tables(QLRNum + 2).Cell(FWLY + 1, 5).Range.Start, _
        wordDoc.Tables(QLRNum + 2).Cell(FWLY + 1, 9).Range.End)
    range32.Font.Size = 11
    range32.Font.Name = "����"
    
'    maxCSIndex = YBH + 2
'    maxCS = Left(wordDoc.Tables(QLRNum + 2).Cell(maxCSIndex, 5).Range.Text, Len(wordDoc.Tables(QLRNum + 2).Cell(maxCSIndex, 5).Range.Text) - 2)
'    For i = YBH + 3 To FWLY - 6
'        cCS = Left(wordDoc.Tables(QLRNum + 2).Cell(i, 5).Range.Text, Len(wordDoc.Tables(QLRNum + 2).Cell(i, 5).Range.Text) - 2)
'        If cCS <> "" And cCS > maxCS Then
'            wordDoc.Tables(QLRNum + 2).Cell(maxCSIndex, 5).Range.Text = ""
'            maxCSIndex = i
'            maxCS = cCS
'        ElseIf cCS <> "" And cCS <= maxCS Then
'            wordDoc.Tables(QLRNum + 2).Cell(i, 5).Range.Text = ""
'        End If
'    Next
'
'    For i = FWLY - 6 To YBH + 2 Step -1
'        If wordDoc.Tables(QLRNum + 2).Cell(i, 5).Range.Text = Chr(13) + "" Then
'            wordDoc.Tables(QLRNum + 2).Cell(i, 5).Select
'            wordApp.Selection.Rows.Delete
'        End If
'    Next
    
    '''2���������������
    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "2�����������������"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute

    If QLRNum = 1 Then
        wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=2
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-4014, Unicode:=True
        wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=6
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="����", CharacterNumber:=9633, Unicode:=True
    Else
        wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=2
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="����", CharacterNumber:=9633, Unicode:=True
        wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=6
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-4014, Unicode:=True
    End If
    
    '''3����������������
    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "3����������������"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute
    wordApp.Selection.EndKey Unit:=wdRow, Extend:=wdExtend
    BZ3 = wordApp.Selection.Cells(1).RowIndex
    chrBZ3 = wordApp.Selection.Characters.Count
    If chrBZ3 < 174 Then
        p3.IncrementTop -15
    Else
        'p3.IncrementTop 5
    End If
    
    '''4���Ƿ�ռ�ø���
    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "4���Ƿ�ռ�ø��أ�"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute
    
    wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=1
    gengdi = wordApp.Selection.Text
    If gengdi <> "��" Then    '��
        wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=9
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="����", CharacterNumber:=9633, Unicode:=True
        wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=6
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-4014, Unicode:=True
    End If
    
    '''5��լ����ʹ��Ȩ�������
    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "5��լ����ʹ��Ȩ���������"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute

    If QLRNum = 1 Then
        wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=2
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-4014, Unicode:=True
        wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=6
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="����", CharacterNumber:=9633, Unicode:=True
    Else
        wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=2
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="����", CharacterNumber:=9633, Unicode:=True
        wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=6
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-4014, Unicode:=True
    End If

    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "�����ߣ�"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute
    wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=1
    wordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 3
    wordApp.Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    wordApp.Selection.TypeBackspace
    wordApp.Selection.TypeText "                                                          ��  ��  ��"

    '''PPP4��5'''''''''''''''''''''''''''''''''''''''''''''''''''
    '''�����ߡ������
    With wordApp.Selection.Find
        .Text = "�����ߣ�ţ  �"
        .Replacement.Text = "�����ߣ�"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute Replace:=wdReplaceAll
    With wordApp.Selection.Find
        .Text = "����ߣ�������"
        .Replacement.Text = "����ߣ�"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute Replace:=wdReplaceAll
    
    wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=QLRNum + 3
    CurrentPageStart = wordApp.Selection.Start
    CurrentPageEnd = wordApp.Selection.Goto(what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=QLRNum + 4).Start
    Set myRange = wordDoc.Range(CurrentPageStart, CurrentPageEnd)
    If myRange.ShapeRange.Count <> 0 Then
        For Each shp In myRange.ShapeRange
            If shp.WrapFormat.Type = wdWrapBehind Then
                shp.Select
                shp.Delete
            End If
        Next
    End If
    
    wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=QLRNum + 3
    Set p4 = wordDoc.Shapes.AddPicture(Filename:=ZPath + "\p4.png", LinkToFile:=False, SaveWithDocument:=True, Width:=CentimetersToPoints(21), Height:=CentimetersToPoints(29.7), Anchor:=wordApp.Selection.Range) '����ͼƬ
    With p4
        .WrapFormat.Type = wdWrapBehind
        .ZOrder 5
        '.Name = "p4"
        .Select
    End With
    wordApp.Selection.ShapeRange.Align msoAlignLefts, True
    wordApp.Selection.ShapeRange.Align msoAlignTops, True
    wordApp.Selection.ShapeRange.Align msoAlignCenters, True
    
    '''������
    BLC = wordDoc.Tables(QLRNum + 3).Cell(3, 1).Range.Text
    If Not Mid(BLC, 7, 3) Like "2*" Then
        p4.IncrementTop 50
    End If
    
    wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=QLRNum + 4
    CurrentPageStart = wordApp.Selection.Start
    CurrentPageEnd = wordApp.Selection.Goto(what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=QLRNum + 5).Start
    Set myRange = wordDoc.Range(CurrentPageStart, CurrentPageEnd)
    If myRange.ShapeRange.Count <> 0 Then
        myRange.ShapeRange.Select
        myRange.ShapeRange.Delete
    End If
    
    wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=QLRNum + 4
    Set p5 = wordDoc.Shapes.AddPicture(Filename:=ZPath + "\p5.png", LinkToFile:=False, SaveWithDocument:=True, Width:=CentimetersToPoints(21), Height:=CentimetersToPoints(29.7), Anchor:=wordApp.Selection.Range) '����ͼƬ
    With p5
        .WrapFormat.Type = wdWrapBehind
        .ZOrder 5
        '.Name = "p5"
        .Select
    End With
    wordApp.Selection.ShapeRange.Align msoAlignLefts, True
    wordApp.Selection.ShapeRange.Align msoAlignTops, True
    wordApp.Selection.ShapeRange.Align msoAlignCenters, True
    
    '''PPP6'''''''''''''''''''''''''''''''''''''''''''''''''''
    wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=QLRNum + 5
    lastR6 = wordDoc.Tables(tableCount).Rows.Count - 5
    For r6 = 4 To lastR6
        If wordDoc.Tables(tableCount).Cell(r6, 4).Range.Text <> Chr(13) + "" Then
            wordDoc.Tables(tableCount).Cell(r6, 4).Range.Text = ""
        Else
            Exit For
        End If
    Next r6
    
    '''PPP7'''''''''''''''''''''''''''''''''''''''''''''''''''
    wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=QLRNum + 6
    wordApp.Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    CurrentPageStart = wordApp.Selection.Start
    CurrentPageEnd = wordApp.Selection.End
    wordApp.Selection.MoveLeft Unit:=1, Count:=1
    Set myRange = wordDoc.Range(CurrentPageStart, CurrentPageEnd)
    If myRange.ShapeRange.Count <> 0 Then
        myRange.ShapeRange.Select
        myRange.ShapeRange.Delete
    End If
    
    wordApp.Selection.Goto what:=wdGoToPage, Which:=wdGoToAbsolute, Count:=QLRNum + 6
    Set p7 = wordDoc.Shapes.AddPicture(Filename:=ZPath + "\p7.png", LinkToFile:=False, SaveWithDocument:=True, Width:=CentimetersToPoints(21), Height:=CentimetersToPoints(29.7), Anchor:=wordApp.Selection.Range) '����ͼƬ
    p7.Select
    p7.WrapFormat.Type = wdWrapBehind
    p7.ZOrder 5
    p7.Name = "p7"
    wordApp.Selection.ShapeRange.Align msoAlignLefts, True
    wordApp.Selection.ShapeRange.Align msoAlignTops, True
    wordApp.Selection.ShapeRange.Align msoAlignCenters, True
    
    '''�������
    Set range7 = wordDoc.Range( _
        wordDoc.Tables(QLRNum + 5).Cell(23, 2).Range.Start, _
        wordDoc.Tables(QLRNum + 5).Cell(23, 2).Range.End)
    NFZAngle = Int(Rnd * 135 + 1)
    NFZTop = Int(Rnd * 20 + 50)
    NFZLeft = Int(Rnd * 25 + 50)
    Set NFZ = wordDoc.Shapes.AddPicture(Filename:=ZPath + "\ũ����.png", LinkToFile:=False, SaveWithDocument:=True, Anchor:=range7) '����ͼƬ
    With NFZ
        .IncrementTop NFZTop
        .IncrementLeft NFZLeft
        .Rotation = NFZAngle
        .WrapFormat.Type = wdWrapBehind
        .ZOrder 5
        .Select
        .Name = "pRed 7"
    End With
    
    '''Ȩ��������������
    wordApp.Selection.Find.ClearFormatting
    If hege = "�ϸ�" Then
        wordApp.Selection.Find.Text = "�ϸ�"
    Else
        wordApp.Selection.Find.Text = "���ϸ�"
    End If
    With wordApp.Selection.Find
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute
    
    wordApp.Selection.MoveLeft Unit:=wdCharacter, Count:=1
    wordApp.Selection.TypeBackspace
    wordApp.Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-4014, Unicode:=True

End Sub

Sub BdealDoc07(ByVal wordDoc As Word.Document, mainDir, targetHukouList)

    Dim wordApp As Word.Application
    Set wordApp = GetObject(, "Word.Application")
    Set wordApp = wordDoc.Parent
    
    For Each shp In wordDoc.Shapes
        shp.Delete
    Next
    
    wordDoc.Tables(1).Cell(2, 4).Height = CentimetersToPoints(1.1)
    
    ZPath = mainDir + "\" + Mid(targetHukouList(0), 10, 4)
    Set p8 = wordDoc.Shapes.AddPicture(Filename:=ZPath + "\p8.png", LinkToFile:=False, SaveWithDocument:=True, Width:=CentimetersToPoints(21), Height:=CentimetersToPoints(29.7), Anchor:=wordApp.Selection.Range) '����ͼƬ
    p8.Select
    p8.WrapFormat.Type = wdWrapBehind
    p8.ZOrder 5
    wordApp.Selection.ShapeRange.Align msoAlignLefts, True
    wordApp.Selection.ShapeRange.Align msoAlignTops, True
    wordApp.Selection.ShapeRange.Align msoAlignCenters, True
    p8.IncrementTop 6
    
End Sub

Sub BdealDoc02(ByVal wordDoc As Word.Document, mainDir, targetHukouList, num02, hege)

    Dim wordApp As Word.Application
    Set wordApp = GetObject(, "Word.Application")
    Set wordApp = wordDoc.Parent
    
    For Each shp In wordDoc.Shapes
        shp.Delete
    Next
    
    ZPath = mainDir + "\" + Mid(targetHukouList(0), 10, 4)
    Set p9 = wordDoc.Shapes.AddPicture(Filename:=ZPath + "\p9.png", LinkToFile:=False, SaveWithDocument:=True, Width:=CentimetersToPoints(21), Height:=CentimetersToPoints(29.7), Anchor:=wordApp.Selection.Range) '����ͼƬ
    p9.Select
    p9.WrapFormat.Type = wdWrapBehind
    p9.ZOrder 5
    wordApp.Selection.ShapeRange.Align msoAlignLefts, True
    wordApp.Selection.ShapeRange.Align msoAlignTops, True
    wordApp.Selection.ShapeRange.Align msoAlignCenters, True
    
    '''��׼�ĺš�ԭլ����֤��
    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "��׼�ĺ�"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute
    i = wordApp.Selection.Cells(1).RowIndex
    
    wordDoc.Range(Start:=wordDoc.Tables(1) _
        .Cell(i, 2).Range.Start, End:=wordDoc.Tables(1) _
        .Cell(i + 1, 2).Range.End).Select
    With wordApp.Selection.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 10
        .Scaling = 62
    End With
    
    '''����Ԥ���
    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "Ԥ���"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute
    JC0 = wordApp.Selection.Cells(1).RowIndex

    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "���߱��Ǽ�����"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute
    JC1 = wordApp.Selection.Cells(1).RowIndex
    
    wordDoc.Range(Start:=wordDoc.Tables(1) _
        .Cell(JC0 + 2, 1).Range.Start, End:=wordDoc.Tables(1) _
        .Cell(JC1 - 1, 1).Range.End).Select
    With wordApp.Selection.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 10
        .Scaling = 62
    End With
    
    '''���߱��Ǽ�����
    If IsEmpty(hege) = False Then
        wordApp.Selection.Find.ClearFormatting
        If hege <> "�ϸ�" Then
            With wordApp.Selection.Find
                .Text = hege
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
            End With
            wordApp.Selection.Find.Execute
            
            wordApp.Selection.MoveLeft Unit:=wdCharacter, Count:=1
            wordApp.Selection.TypeBackspace
            wordApp.Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-4014, Unicode:=True
        End If
    End If
    
    '''����
    wordDoc.Range(Start:=wordDoc.Tables(1) _
        .Cell(JC1 + 1, 5).Range.Start, End:=wordDoc.Tables(1) _
        .Cell(JC1 + 1, 9).Range.End).Select
    wordApp.Selection.Font.Size = 10
    wordApp.Selection.Font.Name = "����"
    wordApp.Selection.Rows.HeightRule = wdRowHeightAtLeast
    wordApp.Selection.Rows.Height = CentimetersToPoints(1.83)
    
'    maxCSIndex = JC0 + 2
'    maxCS = Left(wordDoc.Tables(1).Cell(maxCSIndex, 5).Range.Text, Len(wordDoc.Tables(1).Cell(maxCSIndex, 5).Range.Text) - 2)
'    For i = JC0 + 3 To JC1 - 1
'        cCS = Left(wordDoc.Tables(1).Cell(i, 5).Range.Text, Len(wordDoc.Tables(1).Cell(i, 5).Range.Text) - 2)
'        If cCS <> "" And cCS > maxCS Then
'            wordDoc.Tables(1).Cell(maxCSIndex, 5).Range.Text = ""
'            maxCSIndex = i
'            maxCS = cCS
'        ElseIf cCS <> "" And cCS <= maxCS Then
'            wordDoc.Tables(1).Cell(i, 5).Range.Text = ""
'        End If
'    Next
'
'    For i = JC1 - 1 To JC0 + 2 Step -1
'        If wordDoc.Tables(1).Cell(i, 5).Range.Text = Chr(13) + "" Then
'            wordDoc.Tables(1).Cell(i, 5).Select
'            wordApp.Selection.Rows.Delete
'        End If
'    Next
'
'    wordDoc.Tables(1).Cell(JC0 + 2, 11).Select
'    wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=1
'    wordApp.Selection.InsertRows 1
'    wordApp.Selection.EndKey Unit:=wdRow, Extend:=wdExtend
'    wordApp.Selection.MoveRight Unit:=wdCharacter, Count:=1
'    wordApp.Selection.InsertRows 1
    
    '''��������&��ͬ����
    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "��ͬ����"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    wordApp.Selection.Find.Execute
    
    If num02 > 0 Then
        wordApp.Selection.MoveLeft Unit:=wdCharacter, Count:=1
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-4014, Unicode:=True
        wordApp.Selection.MoveLeft Unit:=wdCharacter, Count:=12
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="����", CharacterNumber:=9633, Unicode:=True
    Else
        wordApp.Selection.MoveLeft Unit:=wdCharacter, Count:=1
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="����", CharacterNumber:=9633, Unicode:=True
        wordApp.Selection.MoveLeft Unit:=wdCharacter, Count:=12
        wordApp.Selection.TypeBackspace
        wordApp.Selection.InsertSymbol Font:="Wingdings 2", CharacterNumber:=-4014, Unicode:=True
    End If
    
    '''ɾ���ڶ�ҳ
    'pageNo = wordDoc.ComputeStatistics(wdStatisticPages)
    wordApp.Selection.Find.ClearFormatting
    With wordApp.Selection.Find
        .Text = "��������Ԫ��ͼ"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    re = wordApp.Selection.Find.Execute
    
    If re Then
        wordApp.Selection.EndKey Unit:=wdStory, Extend:=wdExtend
        wordApp.Selection.TypeBackspace
        With wordApp.Selection.ParagraphFormat
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 1
        End With
    End If
End Sub
