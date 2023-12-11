Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace AddressExtension
    <FormAttribute("13000000", "Business Objects/FrmAddressComponents.b1f")>
    Friend Class FrmAddressComponents
        Inherits SystemFormBase
        Public objform, objformSales, objformUDF As SAPbouiCOM.Form
        Private WithEvents objMatrix As SAPbouiCOM.Matrix
        Public objCombo As SAPbouiCOM.ComboBox
        Public UDFFormID As String
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("Item_0").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()


        End Sub

        Public Sub OnCustomInitialize()
            Try
                Dim FormID As String = ""
                objform = objaddon.objapplication.Forms.GetForm("13000000", 0)
                objMatrix = objform.Items.Item("10000003").Specific
                If objaddon.objapplication.Forms.ActiveForm.Type.ToString = "133" Then
                    Button0.Item.Visible = True
                    objformSales = objaddon.objapplication.Forms.GetForm("133", 0)
                    FormID = objaddon.objapplication.Forms.GetEventForm(objformSales.Type).TypeID.ToString
                    UDFFormID = -FormID
                    If Not objaddon.objapplication.Menus.Item("6913").Checked = True Then
                        objaddon.objapplication.Menus.Item("6913").Activate()
                        objformUDF = objaddon.objapplication.Forms.GetForm(UDFFormID, 1)
                        objCombo = objformUDF.Items.Item("U_Address").Specific
                    End If
                ElseIf objaddon.objapplication.Forms.ActiveForm.Type.ToString = "179" Then
                    Button0.Item.Visible = True
                    objformSales = objaddon.objapplication.Forms.GetForm("179", 0)
                    FormID = objaddon.objapplication.Forms.GetEventForm(objformSales.Type).TypeID.ToString
                    UDFFormID = -FormID
                    If Not objaddon.objapplication.Menus.Item("6913").Checked = True Then
                        objaddon.objapplication.Menus.Item("6913").Activate()
                        objformUDF = objaddon.objapplication.Forms.GetForm(UDFFormID, 0)
                        objCombo = objformUDF.Items.Item("U_Address").Specific
                    End If
                Else
                    Button0.Item.Visible = False
                End If
            Catch ex As Exception
                GC.Collect()
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub
        Public Sub LoadAddressShipTo(ByVal FormID As String)
            Try
                Dim objCombo, objState, objCountry As SAPbouiCOM.ComboBox
                'Dim StrQuery As String
                Dim Street, StreetNo, Block, City, Zipcode, State, Country, County As String
                'Dim objRs As SAPbobsCOM.Recordset
                objform = objaddon.objapplication.Forms.GetForm(FormID, 0)
                objCombo = objform.Items.Item("U_Address").Specific
                objform = objaddon.objapplication.Forms.GetForm("13000000", 0)
                'If objCombo.Selected.Value = "" Then
                '    objaddon.objapplication.StatusBar.SetText("Please Select a Address", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Exit Sub
                'End If
                'objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objState = objMatrix.Columns.Item("10000017").Cells.Item(1).Specific
                objCountry = objMatrix.Columns.Item("10000019").Cells.Item(1).Specific
                'If objaddon.HANA Then
                '    StrQuery = "select Top 1 ""U_StreetS"",""U_StreetNoS"",""U_BlockS"",""U_CityS"",""U_ZipCodeS"",""U_StateS"",""U_CountryS"",""U_CountyS""  from  ""@B2C1""  where  ""Code"" ='" & objCombo.Selected.Value & "';"
                'Else
                '    StrQuery = "select Top 1  U_StreetS,U_StreetNoS,U_BlockS,U_CityS,U_ZipCodeS,U_StateS,U_CountryS,U_CountyS  from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "';"
                'End If
                'objRs.DoQuery(StrQuery)
                Street = objaddon.objglobalmethods.getSingleValue("select Top 1  U_StreetS from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                StreetNo = objaddon.objglobalmethods.getSingleValue("select Top 1  U_StreetNoS from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                Block = objaddon.objglobalmethods.getSingleValue("select Top 1  U_BlockS from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                City = objaddon.objglobalmethods.getSingleValue("select Top 1  U_CityS from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                Zipcode = objaddon.objglobalmethods.getSingleValue("select Top 1  U_ZipCodeS from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                State = objaddon.objglobalmethods.getSingleValue("select Top 1  U_StateS from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                Country = objaddon.objglobalmethods.getSingleValue("select Top 1  U_CountryS from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                County = objaddon.objglobalmethods.getSingleValue("select Top 1  U_CountyS from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")

                'objRs = objaddon.objglobalmethods.GetList(StrQuery)
                objMatrix.Columns.Item("10000003").Cells.Item(1).Specific.String = Street 'objRs.Fields.Item("U_StreetS").Value
                objMatrix.Columns.Item("10000005").Cells.Item(1).Specific.String = StreetNo 'objRs.Fields.Item("U_StreetNoS").Value '"Xyz Street"
                objMatrix.Columns.Item("10000009").Cells.Item(1).Specific.String = Block 'objRs.Fields.Item("U_BlockS").Value
                objMatrix.Columns.Item("10000011").Cells.Item(1).Specific.String = City 'objRs.Fields.Item("U_CityS").Value
                objMatrix.Columns.Item("10000013").Cells.Item(1).Specific.String = Zipcode 'objRs.Fields.Item("U_ZipCodeS").Value
                objMatrix.Columns.Item("10000015").Cells.Item(1).Specific.String = County 'objRs.Fields.Item("U_CountyS").Value
                objCountry.Select(Country, SAPbouiCOM.BoSearchKey.psk_ByValue)
                objState.Select(State, SAPbouiCOM.BoSearchKey.psk_ByValue)
                objform = objaddon.objapplication.Forms.GetForm("13000000", 0)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objform.Items.Item("10000001").Click()
                End If
                objaddon.objapplication.StatusBar.SetText("Address Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Public Sub LoadAddressBillTo(ByVal FormID As String)
            Try
                Dim objCombo, objState, objCountry As SAPbouiCOM.ComboBox
                Dim StrQuery As String
                Dim Street, StreetNo, Block, City, Zipcode, State, Country, County As String
                'Dim objRs As SAPbobsCOM.Recordset
                objform = objaddon.objapplication.Forms.GetForm(FormID, 0)
                objCombo = objform.Items.Item("U_Address").Specific
                objform = objaddon.objapplication.Forms.GetForm("13000000", 0)
                'If objCombo.Selected.Value = "" Then
                '    objaddon.objapplication.StatusBar.SetText("Please Select a Address", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Exit Sub
                'End If
                objState = objMatrix.Columns.Item("10000017").Cells.Item(1).Specific
                objCountry = objMatrix.Columns.Item("10000019").Cells.Item(1).Specific
                'If objaddon.HANA Then
                '    StrQuery = "select Top 1  ""U_StreetB"",""U_StreetNoB"",""U_BlockB"",""U_CityB"",""U_ZipCodeB"",""U_StateB"",""U_CountryB"",""U_CountyB""  from  ""@B2C1""  where  ""Code"" ='" & objCombo.Selected.Value & "';"
                'Else
                '    StrQuery = "select Top 1  U_StreetB,U_StreetNoB,U_BlockB,U_CityB,U_ZipCodeB,U_StateB,U_CountryB,U_CountyB  from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "';"
                'End If
                'objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objRs.DoQuery(StrQuery)
                Street = objaddon.objglobalmethods.getSingleValue("select Top 1  U_StreetB from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                StreetNo = objaddon.objglobalmethods.getSingleValue("select Top 1  U_StreetNoB from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                Block = objaddon.objglobalmethods.getSingleValue("select Top 1  U_BlockB from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                City = objaddon.objglobalmethods.getSingleValue("select Top 1  U_CityB from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                Zipcode = objaddon.objglobalmethods.getSingleValue("select Top 1  U_ZipCodeB from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                State = objaddon.objglobalmethods.getSingleValue("select Top 1  U_StateB from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                Country = objaddon.objglobalmethods.getSingleValue("select Top 1  U_CountryB from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")
                County = objaddon.objglobalmethods.getSingleValue("select Top 1  U_CountyB from  [@B2C1]  where  Code ='" & objCombo.Selected.Value & "'")

                objMatrix.Columns.Item("10000003").Cells.Item(1).Specific.String = Street 'objRs.Fields.Item("U_StreetS").Value
                objMatrix.Columns.Item("10000005").Cells.Item(1).Specific.String = StreetNo 'objRs.Fields.Item("U_StreetNoS").Value '"Xyz Street"
                objMatrix.Columns.Item("10000009").Cells.Item(1).Specific.String = Block 'objRs.Fields.Item("U_BlockS").Value
                objMatrix.Columns.Item("10000011").Cells.Item(1).Specific.String = City 'objRs.Fields.Item("U_CityS").Value
                objMatrix.Columns.Item("10000013").Cells.Item(1).Specific.String = Zipcode 'objRs.Fields.Item("U_ZipCodeS").Value
                objMatrix.Columns.Item("10000015").Cells.Item(1).Specific.String = County 'objRs.Fields.Item("U_CountyS").Value
                objCountry.Select(Country, SAPbouiCOM.BoSearchKey.psk_ByValue)
                objState.Select(State, SAPbouiCOM.BoSearchKey.psk_ByValue)
                objform = objaddon.objapplication.Forms.GetForm("13000000", 0)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objform.Items.Item("10000001").Click()
                End If
                objaddon.objapplication.StatusBar.SetText("Address Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter

            Dim ValidCust As String = ""
            Try
                objformUDF = objaddon.objapplication.Forms.GetForm(UDFFormID, 0)
                objCombo = objformUDF.Items.Item("U_Address").Specific

                    If objCombo.Selected.Value = "" Then
                        ' objform.Close()
                        objaddon.objapplication.StatusBar.SetText("Please Select a Address", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    Else
                        'objformNew = objaddon.objapplication.Forms.GetForm(FormID, 0)
                        If objformSales.Items.Item("4").Specific.String = "" Then
                            objaddon.objapplication.StatusBar.SetText("Please Select a Customer", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                    End If
                    If objaddon.HANA Then
                        ValidCust = objaddon.objglobalmethods.getSingleValue("select ""U_BPCatgry"" from OCRD where ""CardCode""='" & objformSales.Items.Item("4").Specific.String & "'")
                    Else
                        ValidCust = objaddon.objglobalmethods.getSingleValue("select U_BPCatgry from OCRD where CardCode='" & objformSales.Items.Item("4").Specific.String & "'")
                    End If

                        If ValidCust = "B2C" Then
                            If AddressType = "0" Then
                                LoadAddressShipTo(UDFFormID)
                            Else
                                LoadAddressBillTo(UDFFormID)
                            End If

                        Else
                            objaddon.objapplication.StatusBar.SetText("Customer is not a B2C Category!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End If
                        AddressType = ""
                    End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub
    End Class
End Namespace
