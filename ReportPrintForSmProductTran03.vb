Public Class ReportPrintForSmProductTran03
  
#Region " �ܼƫŧi "
    Private mPGMID As String = "PrintSmProductTran03" '�{���N��
    Private ReportPrintForSmProductTran03UCO As ReportPrintForSmProductTran03UCO
    Private mSmProductTranDataTable As THS.MES.DataService.MIS.dsMIS.SmProductTranDataTable
    Private mFirstLoad As Boolean = True '�즸���J�X��
    Private mConnectionString_THS As String = ""
    Private mConnectionString_MES As String = ""
    Private mConnectionString_MIS As String = ""
    Private tempI As Integer = 0
    Private pTransTypeCode As String() = {"01", ""}
    Private mReportType As String = ""
#End Region

#Region " FormLoad "
    Private Sub PrintSmProductTran_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '�ˮ֬O�_�w�n�J

        If Not Thread.CurrentPrincipal.Identity.IsAuthenticated Then
            Dim LoginForm As New LoginForm
            If LoginForm.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                MessageBox.Show("�L�v�ϥΥ��t��!", "�t�ΰT��", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.Close()
            End If
        End If
        Me.myAuButtonWin.SetButtonInvisible(THS.Security.AUControl.AuButtonWin.ButtonMode.Save)
        Me.myAuButtonWin.SetButtonInvisible(THS.Security.AUControl.AuButtonWin.ButtonMode.Cancel)
        '��l�� UCO
        Me.Initial_Entity()
        '��l�Ҧ������
        Me.InitialAllControls()
        '�]�w�즸���J�X��
        mFirstLoad = False
    End Sub
#End Region

#Region " ComboBox �ƥ�"
    '�t�ϧO�ܧ�
    Private Sub LookUpEdit_LocationID_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LookUpEdit_LocationID.EditValueChanged
        Try
            If mFirstLoad Then Return
            Me.InitialBindGridInfo()
        Catch ex As Exception
            MessageBox.Show("�t�ϧO�ܧ�e����l���ѡI" & vbCrLf & ex.Message, "���~�T��", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    '���~���O�ܧ�
    Private Sub LookUpEdit_ProductType_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LookUpEdit_ProductType.EditValueChanged
        Try
            If mFirstLoad Then Return
            Me.InitialBindGridInfo()
        Catch ex As Exception
            MessageBox.Show("���~���O�ܧ�e����l���ѡI" & vbCrLf & ex.Message, "���~�T��", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    '�����_���ܧ�
    Private Sub LookUpEdit_DeptID_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LookUpEdit_DeptIDS.EditValueChanged, LookUpEdit_DeptIDE.EditValueChanged
        Try
            If mFirstLoad Then Return
            If Me.LookUpEdit_DeptIDS.EditValue > Me.LookUpEdit_DeptIDE.EditValue Then
                MessageBox.Show("�����_�����~�I", "���~�T��", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.InitialBindGridInfo()
        Catch ex As Exception
            MessageBox.Show("�����_���ܧ�e����l���ѡI" & vbCrLf & ex.Message, "���~�T��", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    '����_���ܧ�
    Private Sub DateEdit_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateEdit_S.EditValueChanged, DateEdit_E.EditValueChanged
        Try
            If mFirstLoad Then Return
            If Me.DateEdit_S.DateTime > Me.DateEdit_E.DateTime Then
                MessageBox.Show("����_�����~�I", "���~�T��", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.InitialBindGridInfo()
        Catch ex As Exception
            MessageBox.Show("����_���ܧ�e����l���ѡI" & vbCrLf & ex.Message, "���~�T��", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    'RadioButton�ܧ�
    Private Sub RadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnStockOut.CheckedChanged, rbtnStockInMonth.CheckedChanged, rbtnStockIn.CheckedChanged, rbtnStockReturn.CheckedChanged
        Try
            If Me.rbtnStockIn.Checked = True Then '�J�w��
                mReportType = "1"
                pTransTypeCode(0) = "01"
                pTransTypeCode(1) = ""
            ElseIf Me.rbtnStockInMonth.Checked = True Then '�����
                mReportType = "9"
                pTransTypeCode(0) = "01"
                pTransTypeCode(1) = ""
            ElseIf Me.rbtnStockOut.Checked Then '�X�w��
                mReportType = "2"
                pTransTypeCode(0) = "19"
                pTransTypeCode(1) = ""
            Else '��h�J�x�q��
                mReportType = "3"
                pTransTypeCode(0) = "01"
                pTransTypeCode(1) = "19"
            End If
            Me.InitialBindGridInfo()
        Catch ex As Exception
            MessageBox.Show("RadioButton�ܧ�e����l���ѡI" & vbCrLf & ex.Message, "���~�T��", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
 
#End Region


#Region " Button �ƥ� "
    Private Sub myAuButtonWin_ButtonPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles myAuButtonWin.ButtonPrint_Click
        Dim tempDateS, tempDateE As String
        Dim tempDate As Date
        tempDateS = Format(Year(Me.DateEdit_S.DateTime), "0000") & "/" & Format(Month(Me.DateEdit_S.DateTime), "00") & "/" & "01"
        'tempDate = Format(Year(Me.DateEdit_S.DateTime), "0000") & "/" & Format(Month(Me.DateEdit_S.DateTime) + 1, "00") & "/" & "01"
        tempDate = Format(Year(Me.DateEdit_S.DateTime), "0000") & "/" & Format(Month(Me.DateEdit_S.DateTime), "00") & "/" & "01"
        tempDateE = (tempDate.AddMonths(1)).AddDays(-1)
        'tempDateE = tempDate.AddDays(-1)
        Dim SmProductTranReport As New THS.MES.Report.MISReport.MISReport(mConnectionString_MIS, mConnectionString_THS)
        If mReportType = "9" Then
            '�����
            Dim Report As XtraReport = SmProductTranReport.GetSmProductTranReport07(Me.LookUpEdit_LocationID.EditValue, tempDateS, tempDateE, _
                  Me.LookUpEdit_DeptIDS.EditValue, Me.LookUpEdit_DeptIDE.EditValue, Me.DateEditPrint.DateTime)
            Report.ShowPreviewDialog()
        Else
            Dim Report As XtraReport = SmProductTranReport.GetSmProductTranReport03(Me.LookUpEdit_LocationID.EditValue, Me.DateEdit_S.DateTime, Me.DateEdit_E.DateTime, _
            Me.LookUpEdit_DeptIDS.EditValue, Me.LookUpEdit_DeptIDE.EditValue, Me.DateEditPrint.DateTime, Me.mReportType, mSmProductTranDataTable)
            Report.ShowPreviewDialog()
        End If

    End Sub
#End Region

#Region " �����l�� "
    Private Sub InitialAllControls()
        Try
            '��l�t�ϧO���
            Me.InitialLocation()
            '��l���~���O
            Me.InitialProductType()
            '��l�����O
            Me.InitialDeptID()
            '��l���ʤ���_��
            Me.InitialTransDate()
            '��l�L����
            Me.DateEditPrint.DateTime = Today
            '��l�L��ﶵ
            Me.rbtnStockIn.Checked = True
            '��lGridView
            Me.InitialBindGridInfo()

        Catch ex As Exception
            Throw New Exception("�����l���ѡI" & vbCrLf & ex.Message)
        End Try
    End Sub

    '��l�t�ϧO
    Private Sub InitialLocation()
        Try
            With Me.LookUpEdit_LocationID
                .Properties.DataSource = Me.ReportPrintForSmProductTran03UCO.QueryLocation()
                .Properties.DisplayMember = "LocationDesc"
                .Properties.ValueMember = "LocationID"
                '�]�w�w�]��
                For tempI = 0 To Me.ReportPrintForSmProductTran03UCO.QueryLocation().Rows.Count - 1
                    If Trim(Me.ReportPrintForSmProductTran03UCO.QueryLocation().Rows(tempI).Item("LocationID")) = Trim(My.Settings.DefaultLocation) Then
                        .ItemIndex = tempI
                        Exit For
                    End If
                Next
            End With
            Me.LookUpEdit_LocationID.Enabled = False
        Catch ex As Exception
            Throw New Exception("�t�ϧO��l���ѡI" & vbCrLf & ex.Message)
        End Try

    End Sub
    '��l���~���O
    Private Sub InitialProductType()
        Try
            Dim tempProductTypeForComboDataTable As DataTable = ReportPrintForSmProductTran03UCO.QueryProductTypeForCombo(Me.LookUpEdit_LocationID.EditValue, "*ALL")

            Dim mRow As DataRow = tempProductTypeForComboDataTable.NewRow
            mRow.Item("ProductType") = "*ALL"
            mRow.Item("ProductTypeDesc") = "����"
            tempProductTypeForComboDataTable.Rows.InsertAt(mRow, 0)

            With Me.LookUpEdit_ProductType.Properties
                .DataSource = tempProductTypeForComboDataTable
                .ValueMember = "ProductType"
                .DisplayMember = "ProductTypeDesc"
            End With
            'Me.LookUpEdit_ProductType.ItemIndex = 0
            Me.LookUpEdit_ProductType.EditValue = "*ALL"

        Catch ex As Exception
            Throw New Exception("���~���O��l���ѡI" & vbCrLf & ex.Message)
        End Try
    End Sub
    '��l�����O
    Private Sub InitialDeptID()
        Try
            Dim tempDeptIDForComboDataTable As DataTable = ReportPrintForSmProductTran03UCO.QueryDeptID(Me.LookUpEdit_LocationID.EditValue)

            Dim mRow As DataRow = tempDeptIDForComboDataTable.NewRow
            mRow.Item("DeptID") = "*ALL"
            mRow.Item("DeptDesc") = "����"
            tempDeptIDForComboDataTable.Rows.InsertAt(mRow, 0)

            With Me.LookUpEdit_DeptIDS.Properties
                .DataSource = tempDeptIDForComboDataTable
                .ValueMember = "DeptID"
                .DisplayMember = "DeptDesc"
            End With
            Me.LookUpEdit_DeptIDS.ItemIndex = 0
            Me.LookUpEdit_DeptIDS.EditValue = "*ALL"

            With Me.LookUpEdit_DeptIDE.Properties
                .DataSource = tempDeptIDForComboDataTable
                .ValueMember = "DeptID"
                .DisplayMember = "DeptDesc"
            End With
            Me.LookUpEdit_DeptIDE.ItemIndex = 0
            Me.LookUpEdit_DeptIDE.EditValue = "*ALL"

        Catch ex As Exception
            Throw New Exception("�����O��l���ѡI" & vbCrLf & ex.Message)
        End Try
    End Sub
    '��l���ʤ���_��
    Private Sub InitialTransDate()
        Me.DateEdit_S.DateTime = Today
        Me.DateEdit_E.DateTime = Today
    End Sub

#End Region

#Region " InitialBindGridInfo "

    '��lGridControl
    Private Sub InitialBindGridInfo()
        Try

            Me.mSmProductTranDataTable = ReportPrintForSmProductTran03UCO.QueryAllSmProductTran(Me.LookUpEdit_LocationID.EditValue, Me.DateEdit_S.DateTime, Me.DateEdit_E.DateTime, _
                                                        Me.LookUpEdit_DeptIDS.EditValue, Me.LookUpEdit_DeptIDE.EditValue, pTransTypeCode)
            '���o���
            Dim FDataView As DataView = Me.mSmProductTranDataTable.DefaultView
            Dim tempProductType As String = Me.LookUpEdit_ProductType.EditValue
           FDataView.Sort = "ProductType,TransDate,DeptID"
            If tempProductType = "*ALL" Then
                FDataView.RowFilter = ""
            Else
                FDataView.RowFilter = "ProductType='" & tempProductType & "'"
            End If
            Me.gridSmProductTran.DataSource = FDataView
            Me.DataLayoutControl1.DataSource = FDataView
            If FDataView.Count = 0 Then
                MessageBox.Show("�L������ơI", "���~�T��", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Throw New Exception("GridControl���Ū����l���ѡI" & vbCrLf & ex.Message)
        End Try
    End Sub

#End Region

#Region "Initial_Entity"

    Private Sub Initial_Entity()
        mConnectionString_THS = GlobalPara.GetConnectionString(My.Settings.DefaultLocation, DB.THS)
        mConnectionString_MES = GlobalPara.GetConnectionString(My.Settings.DefaultLocation, DB.MES)
        mConnectionString_MIS = GlobalPara.GetConnectionString(My.Settings.DefaultLocation, DB.MIS)
        ReportPrintForSmProductTran03UCO = New ReportPrintForSmProductTran03UCO(mConnectionString_THS, mConnectionString_MES, mConnectionString_MIS)
    End Sub

#End Region


   
 
End Class
