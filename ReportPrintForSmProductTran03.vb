Public Class ReportPrintForSmProductTran03
  
#Region " 變數宣告 "
    Private mPGMID As String = "PrintSmProductTran03" '程式代號
    Private ReportPrintForSmProductTran03UCO As ReportPrintForSmProductTran03UCO
    Private mSmProductTranDataTable As THS.MES.DataService.MIS.dsMIS.SmProductTranDataTable
    Private mFirstLoad As Boolean = True '初次載入旗標
    Private mConnectionString_THS As String = ""
    Private mConnectionString_MES As String = ""
    Private mConnectionString_MIS As String = ""
    Private tempI As Integer = 0
    Private pTransTypeCode As String() = {"01", ""}
    Private mReportType As String = ""
#End Region

#Region " FormLoad "
    Private Sub PrintSmProductTran_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢核是否已登入

        If Not Thread.CurrentPrincipal.Identity.IsAuthenticated Then
            Dim LoginForm As New LoginForm
            If LoginForm.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                MessageBox.Show("無權使用本系統!", "系統訊息", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.Close()
            End If
        End If
        Me.myAuButtonWin.SetButtonInvisible(THS.Security.AUControl.AuButtonWin.ButtonMode.Save)
        Me.myAuButtonWin.SetButtonInvisible(THS.Security.AUControl.AuButtonWin.ButtonMode.Cancel)
        '初始化 UCO
        Me.Initial_Entity()
        '初始所有的控制項
        Me.InitialAllControls()
        '設定初次載入旗標
        mFirstLoad = False
    End Sub
#End Region

#Region " ComboBox 事件"
    '廠區別變更
    Private Sub LookUpEdit_LocationID_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LookUpEdit_LocationID.EditValueChanged
        Try
            If mFirstLoad Then Return
            Me.InitialBindGridInfo()
        Catch ex As Exception
            MessageBox.Show("廠區別變更畫面初始失敗！" & vbCrLf & ex.Message, "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    '產品類別變更
    Private Sub LookUpEdit_ProductType_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LookUpEdit_ProductType.EditValueChanged
        Try
            If mFirstLoad Then Return
            Me.InitialBindGridInfo()
        Catch ex As Exception
            MessageBox.Show("產品類別變更畫面初始失敗！" & vbCrLf & ex.Message, "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    '部門起迄變更
    Private Sub LookUpEdit_DeptID_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LookUpEdit_DeptIDS.EditValueChanged, LookUpEdit_DeptIDE.EditValueChanged
        Try
            If mFirstLoad Then Return
            If Me.LookUpEdit_DeptIDS.EditValue > Me.LookUpEdit_DeptIDE.EditValue Then
                MessageBox.Show("部門起迄錯誤！", "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.InitialBindGridInfo()
        Catch ex As Exception
            MessageBox.Show("部門起迄變更畫面初始失敗！" & vbCrLf & ex.Message, "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    '日期起迄變更
    Private Sub DateEdit_EditValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateEdit_S.EditValueChanged, DateEdit_E.EditValueChanged
        Try
            If mFirstLoad Then Return
            If Me.DateEdit_S.DateTime > Me.DateEdit_E.DateTime Then
                MessageBox.Show("日期起迄錯誤！", "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.InitialBindGridInfo()
        Catch ex As Exception
            MessageBox.Show("日期起迄變更畫面初始失敗！" & vbCrLf & ex.Message, "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    'RadioButton變更
    Private Sub RadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnStockOut.CheckedChanged, rbtnStockInMonth.CheckedChanged, rbtnStockIn.CheckedChanged, rbtnStockReturn.CheckedChanged
        Try
            If Me.rbtnStockIn.Checked = True Then '入庫單
                mReportType = "1"
                pTransTypeCode(0) = "01"
                pTransTypeCode(1) = ""
            ElseIf Me.rbtnStockInMonth.Checked = True Then '月報表
                mReportType = "9"
                pTransTypeCode(0) = "01"
                pTransTypeCode(1) = ""
            ElseIf Me.rbtnStockOut.Checked Then '出庫單
                mReportType = "2"
                pTransTypeCode(0) = "19"
                pTransTypeCode(1) = ""
            Else '剔退入儲量表
                mReportType = "3"
                pTransTypeCode(0) = "01"
                pTransTypeCode(1) = "19"
            End If
            Me.InitialBindGridInfo()
        Catch ex As Exception
            MessageBox.Show("RadioButton變更畫面初始失敗！" & vbCrLf & ex.Message, "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
 
#End Region


#Region " Button 事件 "
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
            '月報表
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

#Region " 控制項初始化 "
    Private Sub InitialAllControls()
        Try
            '初始廠區別選單
            Me.InitialLocation()
            '初始產品類別
            Me.InitialProductType()
            '初始部門別
            Me.InitialDeptID()
            '初始異動日期起迄
            Me.InitialTransDate()
            '初始印表日期
            Me.DateEditPrint.DateTime = Today
            '初始印表選項
            Me.rbtnStockIn.Checked = True
            '初始GridView
            Me.InitialBindGridInfo()

        Catch ex As Exception
            Throw New Exception("控制項初始失敗！" & vbCrLf & ex.Message)
        End Try
    End Sub

    '初始廠區別
    Private Sub InitialLocation()
        Try
            With Me.LookUpEdit_LocationID
                .Properties.DataSource = Me.ReportPrintForSmProductTran03UCO.QueryLocation()
                .Properties.DisplayMember = "LocationDesc"
                .Properties.ValueMember = "LocationID"
                '設定預設值
                For tempI = 0 To Me.ReportPrintForSmProductTran03UCO.QueryLocation().Rows.Count - 1
                    If Trim(Me.ReportPrintForSmProductTran03UCO.QueryLocation().Rows(tempI).Item("LocationID")) = Trim(My.Settings.DefaultLocation) Then
                        .ItemIndex = tempI
                        Exit For
                    End If
                Next
            End With
            Me.LookUpEdit_LocationID.Enabled = False
        Catch ex As Exception
            Throw New Exception("廠區別初始失敗！" & vbCrLf & ex.Message)
        End Try

    End Sub
    '初始產品類別
    Private Sub InitialProductType()
        Try
            Dim tempProductTypeForComboDataTable As DataTable = ReportPrintForSmProductTran03UCO.QueryProductTypeForCombo(Me.LookUpEdit_LocationID.EditValue, "*ALL")

            Dim mRow As DataRow = tempProductTypeForComboDataTable.NewRow
            mRow.Item("ProductType") = "*ALL"
            mRow.Item("ProductTypeDesc") = "全部"
            tempProductTypeForComboDataTable.Rows.InsertAt(mRow, 0)

            With Me.LookUpEdit_ProductType.Properties
                .DataSource = tempProductTypeForComboDataTable
                .ValueMember = "ProductType"
                .DisplayMember = "ProductTypeDesc"
            End With
            'Me.LookUpEdit_ProductType.ItemIndex = 0
            Me.LookUpEdit_ProductType.EditValue = "*ALL"

        Catch ex As Exception
            Throw New Exception("產品類別初始失敗！" & vbCrLf & ex.Message)
        End Try
    End Sub
    '初始部門別
    Private Sub InitialDeptID()
        Try
            Dim tempDeptIDForComboDataTable As DataTable = ReportPrintForSmProductTran03UCO.QueryDeptID(Me.LookUpEdit_LocationID.EditValue)

            Dim mRow As DataRow = tempDeptIDForComboDataTable.NewRow
            mRow.Item("DeptID") = "*ALL"
            mRow.Item("DeptDesc") = "全部"
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
            Throw New Exception("部門別初始失敗！" & vbCrLf & ex.Message)
        End Try
    End Sub
    '初始異動日期起迄
    Private Sub InitialTransDate()
        Me.DateEdit_S.DateTime = Today
        Me.DateEdit_E.DateTime = Today
    End Sub

#End Region

#Region " InitialBindGridInfo "

    '初始GridControl
    Private Sub InitialBindGridInfo()
        Try

            Me.mSmProductTranDataTable = ReportPrintForSmProductTran03UCO.QueryAllSmProductTran(Me.LookUpEdit_LocationID.EditValue, Me.DateEdit_S.DateTime, Me.DateEdit_E.DateTime, _
                                                        Me.LookUpEdit_DeptIDS.EditValue, Me.LookUpEdit_DeptIDE.EditValue, pTransTypeCode)
            '取得資料
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
                MessageBox.Show("無相關資料！", "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Throw New Exception("GridControl資料讀取初始失敗！" & vbCrLf & ex.Message)
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
