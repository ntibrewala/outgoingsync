Imports Sap.Data.Hana
Imports SAPbobsCOM
Imports Newtonsoft.Json.Linq
Imports Bhavya.SAP.Core

Module manualrecocom

    Dim sapSchema As String = "BHAVYA_POLYFILMS_01"
    Dim controlAccount As String = "210001"
    Dim DBSBankGLAccount As String = "230002"

    Dim oCompany As SAPbobsCOM.Company

    Sub Main()

        Dim runId As String = Guid.NewGuid().ToString()

        Logger.StartSession()
        Logger.Log("========== [ManualReco START] RunId=" & runId & " ==========")

        Try
            Dim hasRecords As Boolean = False

            Using roConn = HanaConnectionManager.GetConnection(sapSchema, HanaConnectionManager.HanaConnectionMode.ReadOnlyMode)
                roConn.Open()
                Logger.Log("[" & runId & "] Connected to HANA (ReadOnly)")

                hasRecords = HasUnreconciledPayments(roConn)
                Logger.Log("[" & runId & "] HasUnreconciledPayments = " & hasRecords)
            End Using

            If Not hasRecords Then
                Logger.Log("[" & runId & "] No records found. Exiting.")
                Exit Sub
            End If

            Logger.Log("[" & runId & "] Connecting to SAP DI API...")
            oCompany = SapCompanyManager.Connect(sapSchema)
            Logger.Log("[" & runId & "] SAP Connected")

            Using conn = HanaConnectionManager.GetConnection(sapSchema, HanaConnectionManager.HanaConnectionMode.ReadWriteMode)
                conn.Open()
                Logger.Log("[" & runId & "] Connected to HANA (ReadWrite)")

                ProcessUnreconciledPayments(conn, runId)
            End Using

        Catch ex As Exception
            Logger.Log("[" & runId & "] ERROR: FATAL ERROR: " & ex.ToString())
        Finally
            Logger.Log("[" & runId & "] Disconnecting SAP...")
            SapCompanyManager.Disconnect(oCompany)

            Logger.Log("========== [ManualReco END] RunId=" & runId & " ==========")
            Logger.EndSession()
        End Try

    End Sub

    Sub ProcessUnreconciledPayments(conn As HanaConnection, runId As String)

        Logger.Log("[" & runId & "] Fetching pending payments...")

        Dim sql As String =
        "SELECT P.""Id"", P.""VendorCode"", P.""DocNum"" AS PaymentDocEntry, " &
        "TO_NVARCHAR(P.""InvoicesJson"") AS InvoicesJson, " &
        "IFNULL(P.""Reconciled"",'N') AS ReconState " &
        "FROM ""DBS_BANK"".""PENDING_PAYMENTS"" P " &
        "WHERE P.""Processed"" = 'Y' " &
        "AND IFNULL(P.""Reconciled"",'N') IN ('N','1','2') " &
        "AND P.""InvoicesJson"" IS NOT NULL " &
        "AND TO_NVARCHAR(P.""InvoicesJson"") <> '[]'"

        Using cmd As New HanaCommand(sql, conn)
            Using reader = cmd.ExecuteReader()

                While reader.Read()

                    Dim id = reader("Id").ToString()
                    Dim vendor = reader("VendorCode").ToString()
                    Dim paymentDocEntry = Convert.ToInt32(reader("PaymentDocEntry"))
                    Dim json = reader("InvoicesJson").ToString()
                    Dim state = reader("ReconState").ToString()

                    Logger.Log("[" & runId & "] START Payment | ID=" & id & " | Vendor=" & vendor & " | DocEntry=" & paymentDocEntry & " | State=" & state)
                    Logger.Log("[" & runId & "] JSON = " & json)

                    Try
                        ReconcilePaymentUsingDIAPI(vendor, paymentDocEntry, json, conn, runId)

                        UpdateReconciled(id, "Y", conn)
                        Logger.Log("[" & runId & "] SUCCESS Payment ID=" & id)

                    Catch ex As Exception

                        Dim nextState As String = "1"
                        If state = "1" Then nextState = "2"
                        If state = "2" Then nextState = "F"

                        UpdateReconciled(id, nextState, conn)
                        Logger.Log("[" & runId & "] Updated Reconciled state to " & nextState & " for ID=" & id)
                        Logger.Log("[" & runId & "] ERROR: FAILED Payment ID=" & id & " | NextState=" & nextState & " | Error=" & ex.ToString())

                    End Try

                    Logger.Log("[" & runId & "] END Payment ID=" & id)

                End While

            End Using
        End Using

    End Sub

    Sub ReconcilePaymentUsingDIAPI(vendor As String, paymentDocEntry As Integer, invoiceJson As String, conn As HanaConnection, runId As String)

        Logger.Log("[" & runId & "] Reconciliation START | PaymentDocEntry=" & paymentDocEntry)

        Dim arr = JArray.Parse(invoiceJson)

        Dim paymentTransId As Integer = 0
        Dim paymentLineId As Integer = 0

        GetPaymentJEInfo(paymentDocEntry, paymentTransId, paymentLineId, conn, runId)

        Logger.Log("[" & runId & "] Payment JE | TransId=" & paymentTransId & " | LineId=" & paymentLineId)

        Dim oCmpSrv = oCompany.GetCompanyService()
        Dim oReconSrv = oCmpSrv.GetBusinessService(ServiceTypes.InternalReconciliationsService)

        Dim oRecon = oReconSrv.GetDataInterface(
        InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)

        oRecon.ReconDate = DateTime.Now
        oRecon.CardOrAccount = CardOrAccountEnum.coaCard

        Dim total As Double = 0

        For Each item In arr

            Dim invTransId = CInt(item("InvoiceNumber"))

            Dim docEntry As Integer = 0
            Dim lineId As Integer = 0
            Dim amount As Double = 0

            GetInvoiceJEInfo(invTransId, docEntry, lineId, amount, conn, runId)

            Logger.Log("[" & runId & "] Invoice | TransId=" & invTransId & " | LineId=" & lineId & " | Amount=" & amount)

            Dim row = oRecon.InternalReconciliationOpenTransRows.Add()
            row.Selected = BoYesNoEnum.tYES
            row.TransId = invTransId
            row.TransRowId = lineId
            row.ReconcileAmount = amount

            total += amount

        Next

        Logger.Log("[" & runId & "] Total Invoice Amount = " & total)

        Dim payRow = oRecon.InternalReconciliationOpenTransRows.Add()
        payRow.Selected = BoYesNoEnum.tYES
        payRow.TransId = paymentTransId
        payRow.TransRowId = paymentLineId
        payRow.ReconcileAmount = total

        Logger.Log("[" & runId & "] Posting reconciliation...")

        oReconSrv.Add(oRecon)

        Logger.Log("[" & runId & "] Reconciliation DONE | PaymentDocEntry=" & paymentDocEntry)

    End Sub

    Sub UpdateReconciled(id As String, state As String, conn As HanaConnection)

        Using cmd As New HanaCommand(
        "UPDATE ""DBS_BANK"".""PENDING_PAYMENTS"" SET ""Reconciled""=? WHERE ""Id""=?", conn)

            cmd.Parameters.AddWithValue("p_state", state)
            cmd.Parameters.AddWithValue("p_id", id)
            cmd.ExecuteNonQuery()

        End Using

    End Sub

    Function HasUnreconciledPayments(conn As HanaConnection) As Boolean

        Using cmd As New HanaCommand(
        "SELECT 1 FROM ""DBS_BANK"".""PENDING_PAYMENTS"" " &
        "WHERE ""Processed""='Y' AND IFNULL(""Reconciled"",'N') IN ('N','1','2') LIMIT 1", conn)

            Return cmd.ExecuteScalar() IsNot Nothing

        End Using

    End Function

    Private Sub GetInvoiceJEInfo(invTransId As Integer,
                            ByRef docEntry As Integer,
                            ByRef lineId As Integer,
                            ByRef actualAmount As Double,
                            conn As HanaConnection,
                            runId As String)

        Using cmd As New HanaCommand(
            "SELECT P.""DocEntry"", J.""Line_ID"", J.""Credit"" " &
            "FROM """ & sapSchema & """.""OPCH"" P " &
            "INNER JOIN """ & sapSchema & """.""JDT1"" J ON P.""TransId"" = J.""TransId"" " &
            "WHERE P.""TransId"" = " & invTransId &
            " AND J.""Account"" = '" & controlAccount & "' AND J.""Credit"" > 0", conn)

            Using reader = cmd.ExecuteReader()
                If reader.Read() Then
                    docEntry = Convert.ToInt32(reader("DocEntry"))
                    lineId = Convert.ToInt32(reader("Line_ID"))
                    actualAmount = Convert.ToDouble(reader("Credit"))
                Else
                    Throw New Exception("Invoice JE not found | TransId=" & invTransId)
                End If
            End Using

        End Using

    End Sub

    Private Sub GetPaymentJEInfo(paymentDocEntry As Integer,
                            ByRef transId As Integer,
                            ByRef lineId As Integer,
                            conn As HanaConnection,
                            runId As String)

        Using cmd As New HanaCommand(
            "SELECT V.""TransId"", J.""Line_ID"" " &
            "FROM """ & sapSchema & """.""OVPM"" V " &
            "INNER JOIN """ & sapSchema & """.""JDT1"" J ON V.""TransId"" = J.""TransId"" " &
            "WHERE V.""DocEntry"" = " & paymentDocEntry &
            " AND J.""Debit"" > 0 AND J.""Account"" <> '" & DBSBankGLAccount & "'", conn)

            Using reader = cmd.ExecuteReader()
                If reader.Read() Then
                    transId = Convert.ToInt32(reader("TransId"))
                    lineId = Convert.ToInt32(reader("Line_ID"))
                Else
                    Throw New Exception("Payment JE not found | DocEntry=" & paymentDocEntry)
                End If
            End Using

        End Using

    End Sub
    Public Sub RunManualReco()
        Main()
    End Sub
End Module