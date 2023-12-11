Imports SAPbouiCOM
Namespace AddressExtension

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods

        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    'Case "150"
                    '    ItemMaster_MenuEvent(pVal, BubbleEvent)
                    'Case "SUBBOM"
                    'SubContractingBOM_MenuEvent(pVal, BubbleEvent)

                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByVal BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                Else
                    Select Case pval.MenuUID
                        Case "1281"
                        Case Else
                    End Select
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "SubContractingPO"

        Private Sub SubContractingPO_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0, Matrix2, Matrix4, Matrix3 As SAPbouiCOM.Matrix
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("mtxoutput").Specific
                Matrix2 = objform.Items.Item("Mtxinput").Specific
                Matrix4 = objform.Items.Item("MtxCosting").Specific
                Matrix3 = objform.Items.Item("mtxoutput").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"
                            BubbleEvent = False
                        Case "1292"

                    End Select
                Else
                    Select Case pval.MenuUID

                        Case "1281" 'Find Mode
                            objform.Items.Item("txtdocnum").Enabled = True
                            objform.Items.Item("txtcode").Enabled = True
                            objform.Items.Item("txtctper").Enabled = True
                            objform.Items.Item("txtsitem").Enabled = True
                            objform.Items.Item("docdate").Enabled = True
                            objform.Items.Item("deldate").Enabled = True
                            objform.Items.Item("txtentry").Enabled = True
                            'objform.ActiveItem = "txtdocnum"
                            objform.Items.Item("btngendoc").Enabled = True
                            objform.Items.Item("BtnView").Enabled = True
                            objform.Items.Item("BtnInv").Enabled = True
                            objform.Items.Item("BtnGIssue").Enabled = True
                            objform.Items.Item("btnload").Enabled = True
                            objform.Items.Item("btnOutput").Enabled = True
                            objform.Items.Item("BtnScrap").Enabled = True
                        Case "1282" ' Add Mode
                            objform.Items.Item("btngendoc").Enabled = False
                            objform.Items.Item("btnload").Enabled = False
                            objform.Items.Item("BtnView").Enabled = False
                            objform.Items.Item("BtnInv").Enabled = False
                            objform.Items.Item("BtnGIssue").Enabled = False
                            objform.Items.Item("btnOutput").Enabled = False
                            objform.Items.Item("BtnScrap").Enabled = False
                            objform.Items.Item("posdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("docdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix4, "Code", "#")
                            objform.Items.Item("txtdocnum").Specific.string = objaddon.objglobalmethods.GetNextDocNum_Value("@MIPL_OPOR")
                            objform.Items.Item("txtentry").Specific.string = objaddon.objglobalmethods.GetNextDocEntry_Value("@MIPL_OPOR")
                        Case "1288", "1289", "1290", "1291"
                            For j = 1 To Matrix4.RowCount
                                If Matrix4.Columns.Item("Stat").Cells.Item(j).Specific.String = "C" Then
                                    Matrix4.CommonSetting.SetRowEditable(j, False)
                                End If
                            Next
                            For j = 1 To Matrix3.RowCount
                                If Matrix3.Columns.Item("Colstat").Cells.Item(j).Specific.String = "C" Then
                                    Matrix3.CommonSetting.SetRowEditable(j, False)
                                End If
                            Next
                        Case "1292"

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#End Region

        

        Private Sub ItemMaster_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim objCheck As SAPbouiCOM.CheckBox
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                objCheck = objform.Items.Item("Global").Specific

                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel

                            
                        Case "1293"

                        Case "1292"

                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode  
                        Case "1282" ' Add Mode
                            

                            'objform.Items.Item("txtentry").Specific.string = objaddon.objglobalmethods.getSingleValue("select Count(*)+1 ""DocEntry"" from ""@MIPL_OBOM""")
                        Case "1288", "1289", "1290", "1291"
                           
                            'objaddon.objapplication.Menus.Item("1300").Activate()
                        Case "1293"
                            'If objform.Mode = BoFormMode.fm_OK_MODE Then objform.Mode = BoFormMode.fm_UPDATE_MODE
                            'objform.Update()
                            'objform.Refresh()
                        Case "1292"

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub



    End Class
End Namespace