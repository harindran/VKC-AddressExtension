﻿Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace AddressExtension
    <FormAttribute("179", "Business Objects/FrmSalesReturn.b1f")>
    Friend Class FrmSalesReturn
        Inherits SystemFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("10002101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("10002102").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("179", 0)
            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            'Ship To
            Try
                AddressType = objaddon.objglobalmethods.AddressType("0")
            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents Button1 As SAPbouiCOM.Button

        Private Sub Button1_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button1.ClickAfter
            'Bill To
            Try
                AddressType = objaddon.objglobalmethods.AddressType("1")
            Catch ex As Exception

            End Try
        End Sub
    End Class
End Namespace
