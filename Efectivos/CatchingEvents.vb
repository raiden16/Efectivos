Public Class CatchingEvents

    Friend WithEvents SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Friend SBOCompany As SAPbobsCOM.Company '//OBJETO COMPAÑIA
    Friend csDirectory As String '//DIRECTORIO DONDE SE ENCUENTRAN LOS .SRF
    Dim TransId, Comments, Cuenta, DocDate, JDTNUM As String
    Dim DocTotal As Double

    Public Sub New()
        MyBase.New()
        SetAplication()
        SetConnectionContext()
        ConnectSBOCompany()

        setFilters()

    End Sub

    '//----- ESTABLECE LA COMUNICACION CON SBO
    Private Sub SetAplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        Try
            SboGuiApi = New SAPbouiCOM.SboGuiApi
            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sConnectionString)
            SBOApplication = SboGuiApi.GetApplication()
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la aplicación SBO " & ex.Message)
            System.Windows.Forms.Application.Exit()
            End
        End Try
    End Sub

    '//----- ESTABLECE EL CONTEXTO DE LA APLICACION
    Private Sub SetConnectionContext()
        Try
            SBOCompany = SBOApplication.Company.GetDICompany
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con el DI")
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End
            'Finally
        End Try
    End Sub

    '//----- CONEXION CON LA BASE DE DATOS
    Private Sub ConnectSBOCompany()
        Dim loRecSet As SAPbobsCOM.Recordset
        Try
            '//ESTABLECE LA CONEXION A LA COMPAÑIA
            csDirectory = My.Application.Info.DirectoryPath
            If (csDirectory = "") Then
                System.Windows.Forms.Application.Exit()
                End
            End If
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la BD. " & ex.Message)
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End
        Finally
            loRecSet = Nothing
        End Try
    End Sub


    '//----- ESTABLECE FILTROS DE EVENTOS DE LA APLICACION
    Private Sub setFilters()
        Dim lofilter As SAPbouiCOM.EventFilter
        Dim lofilters As SAPbouiCOM.EventFilters

        Try

            lofilters = New SAPbouiCOM.EventFilters
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            lofilter.AddEx(392) '////// FORMA REGISTRO EN EL DIARIO

            SBOApplication.SetFilter(lofilters)

        Catch ex As Exception
            SBOApplication.MessageBox("SetFilter: " & ex.Message)
        End Try

    End Sub


    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ''// METODOS PARA EVENTOS DE LA APLICACION
    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select

    End Sub

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// METODOS PARA MANEJO DE EVENTOS ITEM
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SBOApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent

        Try

            ''SBOApplication.MessageBox("Action: " & pVal.Before_Action & "  Type: " & pVal.FormTypeEx)
            If pVal.Before_Action = True And pVal.FormTypeEx <> "" Then

                Select Case pVal.FormTypeEx
                    '////////////////FORMA PARA ACTIVAR LICENCIA
                    Case 392
                        frmPOControllerBefore(FormUID, pVal)
                End Select
            End If

            If pVal.Before_Action = False And pVal.FormTypeEx <> "" Then

                Select Case pVal.FormTypeEx
                    '////////////////FORMA PARA ACTIVAR LICENCIA
                    Case 392
                        frmPOControllerAfter(FormUID, pVal)
                End Select
            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error SBOApplication_ItemEvent: " & ex.Message)

        End Try

    End Sub


    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS FORMA PEDIDOS DE COMPRAS
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub frmPOControllerBefore(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)
        'Dim oPO As PO
        'Dim otekPagos As FrmtekPagos
        Dim coForm As SAPbouiCOM.Form
        Dim stTabla As String
        Dim oDatatable As SAPbouiCOM.DBDataSource

        Try

            Select Case pVal.EventType

                                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID
                                    '--- Boton Movimientos del Pedido
                        Case 1

                            stTabla = "OJDT"
                            coForm = SBOApplication.Forms.Item(FormUID)

                            oDatatable = coForm.DataSources.DBDataSources.Item(stTabla)
                            TransId = RTrim(oDatatable.GetValue("Project", 0))
                            Comments = RTrim(oDatatable.GetValue("Memo", 0))
                            DocDate = oDatatable.GetValue("RefDate", 0)
                            'JDTNUM = oDatatable.GetValue("JDT_NUM", 0)

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Pedido de Compras. " & ex.Message)
        Finally
            'oPO = Nothing
        End Try
    End Sub

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS FORMA PEDIDOS DE COMPRAS
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub frmPOControllerAfter(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)
        'Dim oPO As PO
        'Dim otekPagos As FrmtekPagos
        Dim oOBNK As OBNK
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim vJE As SAPbobsCOM.JournalEntries
        vJE = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

        Try

            Select Case pVal.EventType

                                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID
                                    '--- Boton Movimientos del Pedido
                        Case 1

                            JDTNUM = vJE.JdtNum.ToString

                            stQueryH = "Select T0.""Account"", T1.""LocTotal"" from JDT1 T0 Inner join OJDT T1 on T1.""TransId""=T0.""TransId"" where T1.""Memo""='" & Comments & "' and T1.""RefDate""='" & DocDate & "' and T0.""Project""='" & TransId & "' and T0.""Line_ID""<>0 group by T0.""Account"", T1.""LocTotal"""
                            oRecSetH.DoQuery(stQueryH)

                            If oRecSetH.RecordCount = 1 Then

                                oRecSetH.MoveFirst()
                                Cuenta = oRecSetH.Fields.Item("Account").Value

                                If Cuenta = "110101060" Then

                                    DocTotal = oRecSetH.Fields.Item("LocTotal").Value
                                    oOBNK = New OBNK
                                    oOBNK.UpdateOBNK(Comments, DocTotal)

                                End If

                            End If

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Pedido de Compras. " & ex.Message)
        Finally
            'oPO = Nothing
        End Try
    End Sub

End Class
