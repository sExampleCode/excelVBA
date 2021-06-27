VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJejuCenterPicture 
   Caption         =   "제주CCTV통합관제센터 - 장애처리일지"
   ClientHeight    =   8016
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10620
   OleObjectBlob   =   "frmJejuCenterPicture.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmJejuCenterPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnInsert_Click()
    Dim var As Variant
    Dim rngDB As Range
    
    Set rngDB = Sheet3.Cells(Rows.Count, "a").End(xlUp)
    
    '입력 내용을 배열 변수에 저장
    var = Array(CInt(Me.lblNo), Me.cbDivision, Me.txtPlace, CDate(Me.dtS), Me.txtErrContents, CDate(Me.dtE), Me.txtErrAS, _
                    Me.cbResult, Me.txtEtc, Me.txtPath, Me.lstFile.SelectedItem.SubItems(1))

    With rngDB
        .Offset(1).Resize(1, 11) = var
        
        '열너비자동 테두리
        Call AutoFit_Border(.CurrentRegion)
    End With
    
    '입력값 초기화
    Call Init_InputValue
    
    'lstDB 데이터 가져오기
    Call Get_lstDB
    
    '버튼비활성화 입력수정삭제
    Call BtnAll_UnEnable
End Sub

'입력값 초기화
Sub Init_InputValue()
    Me.lblNo = vbNullString
    Me.txtPlace = vbNullString
    Me.txtErrContents = vbNullString
    Me.txtErrAS = vbNullString
    Me.txtEtc = vbNullString
End Sub
'종료
Private Sub btnQuit_Click()
    Dim lQuit As Long
    
    lQuit = MsgBox("저장하시겠습니까?", vbYesNoCancel + vbInformation, "종료")
    
    '예 선택시 저장
    If lQuit = vbYes Then
        ThisWorkbook.Save
        Unload Me
    ElseIf lQuit = vbNo Then
        Unload Me
    End If
End Sub

'버튼비활성화 입력수정삭제
Sub BtnAll_UnEnable()
    Me.btnInsert.Enabled = False
    Me.btnUpdate.Enabled = False
    Me.btnDelete.Enabled = False
End Sub

'버튼활성화 수정삭제
Sub BtnUpdateDelete_Enable()
    Me.btnInsert.Enabled = False
    Me.btnUpdate.Enabled = True
    Me.btnDelete.Enabled = True
End Sub

'버튼활성화 입력
Sub BtnInsert_Enable()
    Me.btnInsert.Enabled = True
    Me.btnUpdate.Enabled = False
    Me.btnDelete.Enabled = False
End Sub

Private Sub btnReset_Click()
    Dim rngDB As Range, rngPICTURE As Range
    Dim ans As Long
    
    Set rngDB = Sheet3.Range("a1").CurrentRegion
    Set rngPICTURE = Sheet4.Range("a1:q13")
    
    '입력 데이터가 있다면
    If rngDB.Rows.Count > 1 Then
        ans = MsgBox("입력된 모든 데이터를 초기화하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "초기화")
        
        If ans = vbYes Then
            With rngDB
                '입력된 모든값 삭제
                .Offset(1).Resize(rngDB.Rows.Count - 1).EntireRow.Delete
            End With
            
            '특정영역 도형삭제
            Call Del_Shps_In_Range(rngPICTURE)
            
            'lstDB 데이터 가져오기
            Call Get_lstDB
            
            '버튼활성화 수정삭제
            Call BtnAll_UnEnable
        End If

    End If
    
    Set rngDB = Nothing
End Sub

Private Sub btnUpdate_Click()
    '찾고 수정
    Call Find_Update
End Sub

Private Sub lstDB_Click()
    Dim li As ListItem
    Dim rngDB As Range, c As Range
    Dim sFilePath As String
    
    If Me.lstDB.ListItems.Count > 0 Then
        Me.lblNo = Me.lstDB.SelectedItem
        Me.cbDivision = Me.lstDB.SelectedItem.SubItems(1)
        Me.txtPlace = Me.lstDB.SelectedItem.SubItems(2)
        Me.dtS = CDate(Me.lstDB.SelectedItem.SubItems(3))
        Me.txtErrContents = Me.lstDB.SelectedItem.SubItems(4)
        Me.dtE = CDate(Me.lstDB.SelectedItem.SubItems(5))
        Me.txtErrAS = Me.lstDB.SelectedItem.SubItems(6)
        Me.cbResult = Me.lstDB.SelectedItem.SubItems(7)
        Me.txtEtc = Me.lstDB.SelectedItem.SubItems(8)
        
        '찾고 사진 넣기
        Call Find_Image
        
        '버튼활성화 수정삭제
        Call BtnUpdateDelete_Enable
    End If
    
    Set rngDB = Nothing
    Set c = Nothing
End Sub

'찾기 사진 넣기
Sub Find_Image()
    Dim rngPICTURE As Range, rngDB As Range
    Dim strAddr As String                         '처음 검색하여 찾은 셀의 주소 넣을 변수
    Dim c As Range                                '검색하여 찾은 영역을 넣을 변수
    Dim r As Range
    Dim shp As Shape
    Dim sFilePath As String
    
    If Me.lstDB.ListItems.Count > 0 Then
        
        Set rngDB = Sheet3.Range("a1", Sheet3.Cells(Rows.Count, "a").End(xlUp))
        
        '입력된 데이터가 1개 이상이라면
        If rngDB.Rows.Count > 1 Then
            With rngDB              '현재시트 사용영역에서
                Set c = .Find(What:=Me.lstDB.SelectedItem, Lookat:=xlWhole) '처음 일치하는 데이터("가")를 찾아 C에 넣고
                
                If Not c Is Nothing Then                  '만일 일치하는 데이터가 있으면
                    strAddr = c.Address                   '첫 일치하는 주소를 strAddr에 넣고
                    
                    Do                                           '다음을 실행
                    
                    Set r = c
                    
                    Loop While Not c Is Nothing And c.Address <> strAddr
                    
                    '일치하는 데이터 없고 처음 일치한 주소 아닐때까지 반복
                End If
            
            End With
        End If
    End If
    
    If Not r Is Nothing Then
    
        Set rngPICTURE = Sheet4.Range("a1:q13")
        
        '특정영역 도형삭제
        Call Del_Shps_In_Range(rngPICTURE)
        
        sFilePath = r.Offset(, 9) & "\" & r.Offset(, 10)
        Set shp = Sheet4.Shapes.AddPicture(sFilePath, msoFalse, msoTrue, rngPICTURE.Left + 2, rngPICTURE.Top + 2, rngPICTURE.Width - 4, rngPICTURE.Height - 4)
    End If
    
    Set rngPICTURE = Nothing
    Set rngDB = Nothing
    Set shp = Nothing
    Set r = Nothing
    Set c = Nothing
    
End Sub

'찾기 수정
Sub Find_Update()

    Dim rngPICTURE As Range, rngDB As Range
    Dim strAddr As String                         '처음 검색하여 찾은 셀의 주소 넣을 변수
    Dim c As Range                                '검색하여 찾은 영역을 넣을 변수
    Dim r As Range
    Dim shp As Shape
    Dim sFilePath As String
    Dim var As Variant
    Dim FSO As FileSystemObject
    Dim f As Folder
    Dim ans As Long
    
    If Me.lstDB.ListItems.Count > 0 Then
        
        Set rngDB = Sheet3.Range("a1", Sheet3.Cells(Rows.Count, "a").End(xlUp))
        
        '입력된 데이터가 1개 이상이라면
        If rngDB.Rows.Count > 1 Then
            With rngDB              '현재시트 사용영역에서
                Set c = .Find(What:=Me.lstDB.SelectedItem, Lookat:=xlWhole) '처음 일치하는 데이터("가")를 찾아 C에 넣고
                
                If Not c Is Nothing Then                  '만일 일치하는 데이터가 있으면
                    strAddr = c.Address                   '첫 일치하는 주소를 strAddr에 넣고
                    
                    Do                                           '다음을 실행
                    
                    Set r = c
                    
                    Loop While Not c Is Nothing And c.Address <> strAddr
                    
                    '일치하는 데이터 없고 처음 일치한 주소 아닐때까지 반복
                End If
            
            End With
        End If
    End If
    
    If Not r Is Nothing Then
        Set rngDB = Sheet3.Cells(Rows.Count, "a").End(xlUp)
        
        '입력 내용을 배열 변수에 저장
        var = Array(CInt(Me.lblNo), Me.cbDivision, Me.txtPlace, CDate(Me.dtS), Me.txtErrContents, CDate(Me.dtE), Me.txtErrAS, _
                        Me.cbResult, Me.txtEtc)
        
        r.Resize(1, 9) = var
        
        With rngDB
            '열너비자동 테두리
            Call AutoFit_Border(.CurrentRegion)
        End With
        
        '입력값 초기화
        Call Init_InputValue
        
        ans = MsgBox("사진을 수정하시겠습니까?", vbYesNo + vbInformation, "수정")
        
        If ans = vbYes Then
            Set rngPICTURE = Sheet4.Range("a1:q13")
        
            '특정영역 도형삭제
            Call Del_Shps_In_Range(rngPICTURE)
            
            With Application.FileDialog(msoFileDialogFilePicker)
                .AllowMultiSelect = False
                .Title = "수정 할 이미지를 선택하세요."
                .Filters.Clear
                .Filters.Add "Images", "*.png; *.jpg; *.jpeg"
                
                If .Show = -1 Then
                    sFilePath = .SelectedItems(1)
                End If
            End With
            
            If sFilePath <> vbNullString Then
                Set FSO = New FileSystemObject
                Set f = FSO.GetFile(sFilePath)
                
                sFilePath = f.ParentFolder & "\" & f.Name
                
                r.Offset(, 9) = f.ParentFolder
                r.Offset(, 10) = f.Name
                
                Set shp = Sheet4.Shapes.AddPicture(sFilePath, msoFalse, msoTrue, rngPICTURE.Left + 2, rngPICTURE.Top + 2, rngPICTURE.Width - 4, rngPICTURE.Height - 4)
            
                
            End If
        
        End If
        
        '버튼비활성화 입력수정삭제
        Call BtnAll_UnEnable
            
        'lstDB 데이터 가져오기
        Call Get_lstDB
        
    End If
    
    Set rngPICTURE = Nothing
    Set rngDB = Nothing
    Set shp = Nothing
    Set r = Nothing
    Set c = Nothing
End Sub

'찾기 삭제
Sub Find_Delete()

    Dim strAddr As String                         '처음 검색하여 찾은 셀의 주소 넣을 변수
    Dim c As Range                                '검색하여 찾은 영역을 넣을 변수
   
    With ActiveSheet.UsedRange              '현재시트 사용영역에서
        Set c = .Find(What:="가", Lookat:=xlWhole) '처음 일치하는 데이터("가")를 찾아 C에 넣고
               
        If Not c Is Nothing Then                  '만일 일치하는 데이터가 있으면
            strAddr = c.Address                   '첫 일치하는 주소를 strAddr에 넣고
           
            Do                                           '다음을 실행
                c.Interior.ColorIndex = 6         '셀의 색을 노란색으로

                Set c = .FindNext(c)             '다음 일치하는 데이터를 찾아 변수에 넣고

            Loop While Not c Is Nothing And c.Address <> strAddr

                              '일치하는 데이터 없고 처음 일치한 주소 아닐때까지 반복
        End If
       
    End With
End Sub

Private Sub lstDB_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Call lstDB_Click
End Sub

Private Sub lstDB_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
    Call lstDB_Click
End Sub

Private Sub lstFile_Click()

    If Me.lstFile.ListItems.Count > 0 Then
        Dim rngPICTURE As Range
        Dim wsPICTURE As Worksheet
        Dim shp As Shape
        Dim sFilePath As String
        
        Set rngPICTURE = Sheet4.Range("a1:q13")
            
        With rngPICTURE
            sFilePath = Me.txtPath & "\" & Me.lstFile.SelectedItem.SubItems(1)
            
            '파일이 있다면
            If sFilePath <> vbNullString Then
                '특정영역 이미지(도형) 삭제
                Call Del_Shps_In_Range(rngPICTURE)
                
                Set wsPICTURE = Sheet4
                
                '사진삽입
                Set shp = wsPICTURE.Shapes.AddPicture(sFilePath, msoFalse, msoCTrue, rngPICTURE.Left + 2, rngPICTURE.Top + 2, rngPICTURE.Width - 4, rngPICTURE.Height - 4)
            End If
        End With
        
         '입력값 초기화
        Call Init_InputValue
        
        '행번호 추가
        Call LastNo
        
        '버튼활성화 입력
        Call BtnInsert_Enable
        
    End If
    
    Set PICTURE = Nothing
    Set wsPICTURE = Nothing
    Set shp = Nothing
    
End Sub

'행번호 추가
Sub LastNo()
    Dim rngDB As Range
    
    Set rngDB = Sheet3.Cells(Rows.Count, "a").End(xlUp)
    
    If IsNumeric(rngDB) Then
        Me.lblNo = rngDB + 1
    Else
        Me.lblNo = 1
    End If
End Sub

'특정영역 이미지(도형) 삭제
Sub Del_Shps_In_Range(rngPICTURE As Range)
  
    Dim shpC As Shape                                  '각각의 도형(shape)을 넣을 변수
    Dim rngShp As Range                               '각 도형의 왼쪽위가 속한 영역을 넣을 변수
    Dim rngAll As Range                                 '삭제할 전체영역을 넣을 변수
  
    Set rngAll = rngPICTURE                                 '선택영역을 삭제할 영역으로 설정
  
    For Each shpC In ActiveSheet.Shapes        '삭제영역내의 각 도형을 순환
        Set rngShp = shpC.TopLeftCell               '각 도형의 왼쪽위지점이 속한 영역을 변수에 넣음
        If Not Intersect(rngAll, rngShp) Is Nothing Then   '도형과 삭제영역이 겹치면
            shpC.Delete                                     '각 도형을 삭제
        End If
    Next shpC
  
    Set rngAll = Nothing                                   '개체변수들 초기화(메모리 비우기)
    Set rngShp = Nothing
   
End Sub

'lstFile KeyDown
Private Sub lstFile_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Call lstFile_Click
End Sub

'lstFile KeyUp
Private Sub lstFile_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
    Call lstFile_Click
End Sub

'그림파일만 lstFile ListView 로 가져오기
Private Sub txtPath_DropButtonClick()
    Dim filePath As String
    Dim FSO As FileSystemObject
    Dim fl As Folder
    Dim f As File
    Dim index As Long
    Dim li As ListItem
    Dim sExtension As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False   '한개의 파일만 선택 가능
        .Title = "이미지 폴더를 선택하세요."
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        End If
    End With
    
    '선택하지 않고 취소한 경우 종료함.
    If filePath = vbNullString Then Exit Sub
    
    '텍스트박스에 폴더경로를 넣음
    Me.txtPath = filePath
    
    Set FSO = New FileSystemObject
    Set fl = FSO.GetFolder(Me.txtPath)
    
    'lstFile 모든 내용 지우기
    Me.lstFile.ListItems.Clear
    
    For Each f In fl.Files
        sExtension = FSO.GetExtensionName(f.Name)
        
        'jpg, jpeg, png 파일만 가능
        If LCase(sExtension) = "jpg" Or LCase(sExtension) = "jpeg" Or LCase(sExtension) = "png" Then
            Set li = Me.lstFile.ListItems.Add(, , index + 1)
            li.ListSubItems.Add , , f.Name
            
            index = index + 1
        End If
    Next
    
    Set FSO = Nothing
    Set fl = Nothing
End Sub

Private Sub UserForm_Initialize()
    Dim rngResult As Range, rngDivision As Range
    
    '경로 열기버튼 스타일 초기화
    With Me.txtPath
        .DropButtonStyle = fmDropButtonStyleEllipsis
        .ShowDropButtonWhen = fmShowDropButtonWhenAlways
    End With
    
    Set rngDivision = Sheet6.Range("a1").CurrentRegion
    
    '1보다 큰경우
    If rngDivision.Rows.Count > 1 Then
        '구분 초기화
        With Me.cbDivision
            .RowSource = rngDivision.Offset(1).Resize(rngDivision.Rows.Count - 1).Address(External:=True)
            .ListIndex = 0
            .Style = fmStyleDropDownList
            .ListStyle = fmListStyleOption
        End With
        
        Call AutoFit_Border(rngDivision)
    End If
    
    '열너비자동, 텍스트 가운데 정렬
    Set rngResult = Sheet7.Range("a1").CurrentRegion
    
    '1보다 큰경우
    If rngResult.Rows.Count > 1 Then
        '조치 결과 초기화
        With Me.cbResult
            .RowSource = rngResult.Offset(1).Resize(rngResult.Rows.Count - 1).Address(External:=True)
            .ListIndex = 0
            .Style = fmStyleDropDownList
            .ListStyle = fmListStyleOption
        End With
        
        '열너비자동, 테두리
        Call AutoFit_Border(rngResult)
    End If
    
    'lstFile 초기화
    With lstFile
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "No", 30
        .ColumnHeaders.Add , , "FileName", 256
        
'        .OLEDropMode = 1         '파일드래그
        .LabelEdit = lvwManual
        .FullRowSelect = True       '전체선택
        .View = lvwReport             '리스트 수정불가
        .CheckBoxes = False        '체크박스 사용여부
        .Gridlines = True               '그리드라인 회색선
    End With
    
    'lstDB 초기화
    With lstDB
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "연번", 30
        .ColumnHeaders.Add , , "구분", 59
        .ColumnHeaders.Add , , "현장명", 59
        .ColumnHeaders.Add , , "장애일자", 59
        .ColumnHeaders.Add , , "장애내용", 59
        .ColumnHeaders.Add , , "조치일자", 59
        .ColumnHeaders.Add , , "조치내용", 59
        .ColumnHeaders.Add , , "조치상황", 59
        .ColumnHeaders.Add , , "특이사항", 59
        
'        .OLEDropMode = 1         '파일드래그
        .LabelEdit = lvwManual
        .FullRowSelect = True       '전체선택
        .View = lvwReport             '리스트 수정불가
        .CheckBoxes = False        '체크박스 사용여부
        .Gridlines = True               '그리드라인 회색선
    End With
    
    'lstDB 데이터 가져오기
    Call Get_lstDB
    
    '초기화
    Me.dtS = Date
    Me.dtE = Date
    Me.lblNo = vbNullString
    Me.txtPath = vbNullString
    Me.txtErrContents = vbNullString
    Me.txtErrAS = vbNullString
    Me.txtEtc = vbNullString
    
    '버튼비활성화
    Call BtnUnenable
    
    Set rngDivision = Nothing
    Set rngResult = Nothing

End Sub

'버튼비활성화
Sub BtnUnenable()
    Me.btnInsert.Enabled = False
    Me.btnUpdate.Enabled = False
    Me.btnDelete.Enabled = False
End Sub

'열너비자동, 테두리
Sub AutoFit_Border(rngStyle As Range)
    Application.ScreenUpdating = False
    
    With rngStyle.CurrentRegion
        .Columns.AutoFit
        .Borders.LineStyle = XlLineStyle.xlLineStyleNone
        .Borders.LineStyle = XlLineStyle.xlContinuous
    End With
    
    Application.ScreenUpdating = True
End Sub

'lstDB 데이터 가져오기
Sub Get_lstDB()
    Dim rngDB As Range
    Dim rowCnt As Long, colCnt As Long
    Dim i As Long, j As Long
    Dim li As ListItem
    
    Set rngDB = Sheet3.Range("a1").CurrentRegion
    
    rowCnt = rngDB.Rows.Count
    colCnt = rngDB.Columns.Count
    
    Me.lstDB.ListItems.Clear
    
    For i = 2 To rowCnt
        Set li = Me.lstDB.ListItems.Add(, , rngDB(i, 1))
        For j = 2 To colCnt
            li.ListSubItems.Add , , rngDB(i, j)
        Next j
    Next i
    
    Set rngDB = Nothing
End Sub

