Public Class OBNK

    Private SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private SBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Dim oInvoice As SAPbobsCOM.Documents

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        SBOApplication = oCatchingEvents.SBOApplication
        SBOCompany = oCatchingEvents.SBOCompany
    End Sub

    Public Function UpdateOBNK(ByVal Account As String, ByVal DocTotal As Double)

        Dim oCuenta As SAPbobsCOM.BankPages
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim BankAcct, Sequence, Fecha, Process As String
        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            Process = Account.Replace(" ", "")
            BankAcct = Process.Substring(0, 9)
            Fecha = Right(Process, 8)

            stQueryH = "Select T0.""Sequence"" from OBNK T0 where T0.""DueDate""='" & Fecha & "' and T0.""CredAmnt""=" & DocTotal & " and T0.""AcctCode""='" & BankAcct & "'"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oRecSetH.MoveFirst()
                Sequence = oRecSetH.Fields.Item("Sequence").Value

                oCuenta = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBankPages)

                oCuenta.GetByKey(BankAcct, Sequence)
                oCuenta.CardCode = "EFECTIVO"
                'oCuenta.ExternalCode = DocEntry
                oCuenta.Update()

            End If


        Catch ex As Exception

            SBOApplication.MessageBox("Error UpdateOBNK: " & ex.Message)

        End Try

    End Function

End Class
