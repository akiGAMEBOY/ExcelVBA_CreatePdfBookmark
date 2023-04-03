Attribute VB_Name = "mdlAcrobat"
Option Explicit
Private Const PDSaveIncremental = &H0
Private Const PDSaveFull = &H1
Private Const PDSaveCopy = &H2
Private Const PDSaveLinearized = &H4
Private Const PDSaveCollectGarbage = &H20
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
'********************************************************************************
'* �������@�bFuncReadBookmark_Acrobat
'* �@�\�@�@�b������f�[�^�ǂݍ���
'*-------------------------------------------------------------------------------
'* �߂�l�@�bLong�i�f�[�^�����j
'* �����@�@�bstrFilepath�F�Ώۃt�@�C��
'* ���ӎ����b�Q�Ɛݒ�ihttp://pdf-file.nnn2.com/?p=204�j
'********************************************************************************
Public Function FuncReadBookmark_Acrobat(strFilepath As String) As Long
    Dim objApp As AcroApp
    Dim objAVDoc As AcroAVDoc
    Dim objPDDoc As New Acrobat.AcroPDDoc
    Dim objPageview As AcroAVPageView
    Dim objJso As Object
    Dim objBkmk As Object
    Dim lngResult As Long
    
    Set objApp = CreateObject("AcroExch.App")
    Set objAVDoc = CreateObject("AcroExch.AVDoc")
    objApp.Hide
    If objPDDoc.Open(strFilepath) = True Then
        If Sheet3.DEBUGMODE Then
            objApp.Show '�A�v���P�[�V�����\��
        End If
        Set objAVDoc = objPDDoc.OpenAVDoc("")
        Set objPageview = objAVDoc.GetAVPageView
        Set objJso = objAVDoc.GetPDDoc.GetJSObject
        Set objBkmk = CallByName(objJso, "bookmarkRoot", VbGet)
        ' �g��\���i������̈ʒu�ɂ��y�[�W�ԍ����Y���Ă��܂��\�������邽�߁j
        lngResult = objPageview.ZoomTo(AVZoomNoVary, 6400)
        Call SubDumpBookmark(objBkmk, objPageview)
        objPDDoc.Close
        objApp.Exit
    End If
    
    Set objBkmk = Nothing
    Set objJso = Nothing
    Set objPageview = Nothing
    Set objAVDoc = Nothing
    Set objPDDoc = Nothing
    Set objApp = Nothing
    
    ' 6�b��~��Acrobat�v���Z�X�Ɍ㏈���̗P�\��^����
    DoEvents
    Sleep 6000
    Call SubTerminateAcrobat ' �v���Z�X���c�����ꍇ�A�����I��
    
    lngResult = Sheet3.dicBookmarkdata.Count
    FuncReadBookmark_Acrobat = lngResult
    
End Function
'********************************************************************************
'* �������@�bSubDumpBookmark
'* �@�\�@�@�b������f�[�^�ǂݍ���(�T�u)
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�bobjBkmk�FbookmarkRoot�I�u�W�F�N�g, objPageview�FAcroAVPageView�I�u�W�F�N�g, lngCol�F�K�w�J�E���g(�C��)
'********************************************************************************
Private Sub SubDumpBookmark(ByVal objBkmk As Object, ByVal objPageview As Object, Optional lngCol As Long = -1)
    Dim lngResult As Long
    Dim cld As Variant, cld2 As Variant
    Dim strMidashi As String
    Dim lngPage As Long
    
    On Error Resume Next
    cld = CallByName(objBkmk, "children", VbGet)
    On Error GoTo 0
    If IsEmpty(cld) = False Then
        lngCol = lngCol + 1
        For Each cld2 In cld
            CallByName cld2, "execute", VbMethod '������I��
            strMidashi = CallByName(cld2, "name", VbGet)
            lngPage = objPageview.GetPageNum + 1
            ' 
            Call Sheet3.SubTempsaveBookmark(strMidashi, lngPage, lngCol)
            SubDumpBookmark cld2, objPageview, lngCol
        Next
        lngCol = lngCol - 1
    End If
End Sub
'********************************************************************************
'* �������@�bFuncInitBookmark
'* �@�\�@�@�b������f�[�^�̏������i���ׂč폜�j
'*-------------------------------------------------------------------------------
'* �߂�l�@�bLong�i0: ����I��, -411: �ُ�I���j
'* �����@�@�bstrFilepath�F�Ώۃt�@�C��
'********************************************************************************
Public Function FuncInitBookmark(strFilepath As String)
    Dim lngResult As Long
    Dim objApp As AcroApp
    Dim objAVDoc As AcroAVDoc
    Dim objPDBookMark As Acrobat.AcroPDBookmark 'INIT
    Dim objPDDoc As New Acrobat.AcroPDDoc
    Dim objPageview As AcroAVPageView
    Dim objJso As Object
    Dim objBkmk As Object
    
    Set objApp = CreateObject("AcroExch.App")
    Set objAVDoc = CreateObject("AcroExch.AVDoc")
    Set objPDBookMark = CreateObject("AcroExch.PDBookmark") 'INIT
    objApp.Hide
    If objPDDoc.Open(strFilepath) = True Then
        If Sheet3.DEBUGMODE Then
            objApp.Show '�A�v���P�[�V�����\��
        End If
        Set objAVDoc = objPDDoc.OpenAVDoc("")
        Set objPageview = objAVDoc.GetAVPageView
        Set objJso = objAVDoc.GetPDDoc.GetJSObject
        Set objBkmk = CallByName(objJso, "bookmarkRoot", VbGet)
        Call SubDeleteBookmark(objBkmk, objPageview, objPDDoc, objPDBookMark)
        lngResult = objPDDoc.Save _
                    (PDSaveFull, _
                     strFilepath)
        If lngResult = -1 Then
            lngResult = 0
        Else
            lngResult = -411
        End If
        objApp.Exit
    End If
    
    Set objBkmk = Nothing
    Set objJso = Nothing
    Set objPDBookMark = Nothing
    Set objPDDoc = Nothing
    Set objPageview = Nothing
    Set objAVDoc = Nothing
    Set objApp = Nothing
    
    ' 6�b��~��Acrobat�v���Z�X�Ɍ㏈���̗P�\��^����
    DoEvents
    Sleep 6000
    Call SubTerminateAcrobat ' �v���Z�X���c�����ꍇ�A�����I��

    FuncInitBookmark = lngResult
    
End Function
'********************************************************************************
'* �������@�bFuncWriteBookmark_Acrobat
'* �@�\�@�@�bPDF�t�@�C���ɂ�����f�[�^��������
'*-------------------------------------------------------------------------------
'* �߂�l�@�bLong�i0: ����I��, -511: �ُ�I���j
'* �����@�@�bstrFilepath�F�Ώۃt�@�C��
'********************************************************************************
Function FuncWriteBookmark_Acrobat(strFilepath As String) As Long
    Dim objApp As AcroApp
    Dim objAVDoc As AcroAVDoc
    Dim objPDDoc As New Acrobat.AcroPDDoc
    Dim objJso As Object
    Dim objBkmkroot As Object
    Dim objBkmk() As Object
    Dim varTemp As Variant
    Dim varBookmarkdata As Variant
    Dim lngCount As Long
    Dim lngDown As Long
    Dim lngResult As Long
    
    Set objApp = CreateObject("AcroExch.App")
    Set objAVDoc = CreateObject("AcroExch.AVDoc")
    objApp.Hide
    If objPDDoc.Open(strFilepath) = True Then
        If Sheet3.DEBUGMODE Then
            objApp.Show '�A�v���P�[�V�����\��
        End If
        Set objAVDoc = objPDDoc.OpenAVDoc("")
        Set objJso = objAVDoc.GetPDDoc.GetJSObject
        ReDim varBookmarkdata(Sheet3.dicBookmarkdata.Count - 1, 5)
        If Not objJso Is Nothing Then
            Set objBkmkroot = objJso.bookmarkRoot
            For lngCount = 0 To Sheet3.dicBookmarkdata.Count - 1
                varBookmarkdata(lngCount, 0) = Sheet3.dicBookmarkdata.item(lngCount)(1) - 1     'Page
                If Sheet3.dicBookmarkdata.item(lngCount)(2) <> "" Then
                    varBookmarkdata(lngCount, 1) = Sheet3.dicBookmarkdata.item(lngCount)(2)     'Midashi
                    varBookmarkdata(lngCount, 2) = lngCount                                     'Index
                    varBookmarkdata(lngCount, 3) = 0                                            'Directory
                ElseIf Sheet3.dicBookmarkdata.item(lngCount)(3) <> "" Then
                    varBookmarkdata(lngCount, 1) = Sheet3.dicBookmarkdata.item(lngCount)(3)
                    varBookmarkdata(lngCount, 2) = lngCount
                    varBookmarkdata(lngCount, 3) = 1
                ElseIf Sheet3.dicBookmarkdata.item(lngCount)(4) <> "" Then
                    varBookmarkdata(lngCount, 1) = Sheet3.dicBookmarkdata.item(lngCount)(4)
                    varBookmarkdata(lngCount, 2) = lngCount
                    varBookmarkdata(lngCount, 3) = 2
                ElseIf Sheet3.dicBookmarkdata.item(lngCount)(5) <> "" Then
                    varBookmarkdata(lngCount, 1) = Sheet3.dicBookmarkdata.item(lngCount)(5)
                    varBookmarkdata(lngCount, 2) = lngCount
                    varBookmarkdata(lngCount, 3) = 3
                ElseIf Sheet3.dicBookmarkdata.item(lngCount)(6) <> "" Then
                    varBookmarkdata(lngCount, 1) = Sheet3.dicBookmarkdata.item(lngCount)(6)
                    varBookmarkdata(lngCount, 2) = lngCount
                    varBookmarkdata(lngCount, 3) = 4
                ElseIf Sheet3.dicBookmarkdata.item(lngCount)(7) <> "" Then
                    varBookmarkdata(lngCount, 1) = Sheet3.dicBookmarkdata.item(lngCount)(7)
                    varBookmarkdata(lngCount, 2) = lngCount
                    varBookmarkdata(lngCount, 3) = 5
                ElseIf Sheet3.dicBookmarkdata.item(lngCount)(8) <> "" Then
                    varBookmarkdata(lngCount, 1) = Sheet3.dicBookmarkdata.item(lngCount)(8)
                    varBookmarkdata(lngCount, 2) = lngCount
                    varBookmarkdata(lngCount, 3) = 6
                ElseIf Sheet3.dicBookmarkdata.item(lngCount)(9) <> "" Then
                    varBookmarkdata(lngCount, 1) = Sheet3.dicBookmarkdata.item(lngCount)(9)
                    varBookmarkdata(lngCount, 2) = lngCount
                    varBookmarkdata(lngCount, 3) = 7
                ElseIf Sheet3.dicBookmarkdata.item(lngCount)(10) <> "" Then
                    varBookmarkdata(lngCount, 1) = Sheet3.dicBookmarkdata.item(lngCount)(10)
                    varBookmarkdata(lngCount, 2) = lngCount
                    varBookmarkdata(lngCount, 3) = 8
                ElseIf Sheet3.dicBookmarkdata.item(lngCount)(11) <> "" Then
                    varBookmarkdata(lngCount, 1) = Sheet3.dicBookmarkdata.item(lngCount)(11)
                    varBookmarkdata(lngCount, 2) = lngCount
                    varBookmarkdata(lngCount, 3) = 9
                End If
                
                ' �K�w����
                If varBookmarkdata(lngCount, 3) = 0 Then                                        ' 0�K�w�Őݒ肳��Ă���ꍇ
                    varBookmarkdata(lngCount, 4) = ""                                           ' Parent
                    varBookmarkdata(lngCount, 5) = 0                                            ' Child Index
                ElseIf varBookmarkdata(lngCount, 3) = varBookmarkdata(lngCount - 1, 3) Then     ' �ЂƂO�̍s�ƊK�w�������ꍇ
                    varBookmarkdata(lngCount, 4) = varBookmarkdata(lngCount - 1, 4)
                    varBookmarkdata(lngCount, 5) = varBookmarkdata(lngCount - 1, 5) + 1
                ElseIf varBookmarkdata(lngCount, 3) > varBookmarkdata(lngCount - 1, 3) Then     ' �K�w���グ��ꍇ
                    varBookmarkdata(lngCount, 4) = varBookmarkdata(lngCount - 1, 2)
                    varBookmarkdata(lngCount, 5) = 0
                ElseIf varBookmarkdata(lngCount, 3) < varBookmarkdata(lngCount - 1, 3) Then     ' �K�w��������ꍇ
                    varBookmarkdata(lngCount, 4) = ""
                    varBookmarkdata(lngCount, 5) = 0
                    ' �O�̍s��1�s������
                    For lngDown = lngCount - 1 To 0 Step -1
                        ' �O�̍s�ɓ����K�w������ꍇ
                        If varBookmarkdata(lngCount, 3) = varBookmarkdata(lngDown, 3) Then
                            varBookmarkdata(lngCount, 4) = varBookmarkdata(lngDown, 4)
                            varBookmarkdata(lngCount, 5) = varBookmarkdata(lngDown, 5) + 1
                            Exit For
                        End If
                    Next
                End If
                
                ' ������쐬
                With objBkmkroot
                     .createChild varBookmarkdata(lngCount, 1), "this.pageNum=" & varBookmarkdata(lngCount, 0), CLng(varBookmarkdata(lngCount, 2))
                End With
            Next
            
            ' �I�u�W�F�N�g��ݒ�
            varTemp = objBkmkroot.Children
            ReDim objBkmk(Sheet3.dicBookmarkdata.Count - 1)
            For lngCount = 0 To UBound(objBkmk)
                On Error Resume Next
                Set objBkmk(lngCount) = varTemp(lngCount)
                On Error GoTo 0
            Next
            
            ' �쐬����������̊K�w�\����ݒ�
            For lngCount = 0 To UBound(varBookmarkdata)
                If varBookmarkdata(lngCount, 4) <> "" Then
                    objBkmk(varBookmarkdata(lngCount, 4)).insertChild objBkmk(varBookmarkdata(lngCount, 2)), varBookmarkdata(lngCount, 5)
                End If
            Next
            
            Set varTemp = Nothing
            ReDim objBkmk(Sheet3.dicBookmarkdata.Count - 1)
            Set objBkmkroot = Nothing
            
        End If
        
        lngResult = objPDDoc.Save(PDSaveIncremental, strFilepath)
        If lngResult = -1 Then
            lngResult = 0
        Else
            lngResult = -511
        End If
        
        objPDDoc.Close
        objApp.Exit
    End If
    
    Set objJso = Nothing
    Set objAVDoc = Nothing
    Set objPDDoc = Nothing
    Set objApp = Nothing
    
    ' 6�b��~��Acrobat�v���Z�X�Ɍ㏈���̗P�\��^����
    DoEvents
    Sleep 6000
    Call SubTerminateAcrobat ' �v���Z�X���c�����ꍇ�A�����I��
    
    FuncWriteBookmark_Acrobat = lngResult
    
End Function
'********************************************************************************
'* �������@�bSubDeleteBookmark
'* �@�\�@�@�b������폜
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�bobjBkmk�FbookmarkRoot�I�u�W�F�N�g, objPageview�FAcroAVPageView�I�u�W�F�N�g, lngCol�F�K�w�J�E���g(�C��)
'********************************************************************************
Private Sub SubDeleteBookmark(ByVal objBkmk As Object, ByVal objPageview As Object, ByVal objPDDoc As Object, ByVal objPDBookMark As Object, Optional lngCol As Long = -1)
    '������̏����o��
    Dim lngResult As Long
    Dim cld As Variant, cld2 As Variant
    Dim strMidashi As String
    Dim lngPage As Long
    
    On Error Resume Next
    cld = CallByName(objBkmk, "children", VbGet)
    On Error GoTo 0
    If IsEmpty(cld) = False Then
        lngCol = lngCol + 1
        For Each cld2 In cld
            CallByName cld2, "execute", VbMethod                        ' ������I��
            strMidashi = CallByName(cld2, "name", VbGet)
            lngPage = objPageview.GetPageNum + 1
            lngResult = objPDBookMark.GetByTitle(objPDDoc, strMidashi)
            If lngResult = True Then
                lngResult = objPDBookMark.Destroy                       ' �폜
            End If
            ' �ċA�Ăяo��
            SubDeleteBookmark cld2, objPageview, objPDDoc, objPDBookMark, lngCol
        Next
        lngCol = 0
    End If
End Sub
'********************************************************************************
'* �������@�bSubTerminateAcrobat
'* �@�\�@�@�bAcrobat�̃v���Z�X�����I��
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�b-
'********************************************************************************
Private Sub SubTerminateAcrobat()
  Dim items As Object
  Dim item As Object
    
  Set items = CreateObject("WbemScripting.SWbemLocator") _
            .ConnectServer.ExecQuery("Select * From Win32_Process Where Name = 'Acrobat.exe'")
  If items.Count > 0 Then
    For Each item In items
      item.Terminate
    Next
  End If
  
  Set items = Nothing
  
End Sub
