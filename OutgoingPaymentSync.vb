Imports Sap.Data.Hana
Imports SAPbobsCOM
Imports Bhavya.SAP.Core

Module OutgoingPaymentSync

    Dim sapSchema As String = "BHAVYA_POLYFILMS_01"
    Dim DBSBankGLAccount As String = "230002"

    Dim oCompany As Company

    '============================================================
    Sub Main()

        Dim runId As String = Guid.NewGuid().ToString()

        Logger.StartSession()
        Logger.Log("========== [START] RunId=" & runId & " ==========")

        Try
            '=========================
            ' HANA CONNECTION
            '=========================
            Using conn = HanaConnectionManager.GetConnection(sapSchema, HanaConnectionManager.HanaConnectionMode.ReadWriteMode)
                conn.Open()
                Logger.Log("[" & runId & "] HANA Connected")

                '-------------------------
                ' Update Bank Status
                '-------------------------
                If HasStatusUpdates(conn) Then
                    Logger.Log("[" & runId & "] Updating bank payment statuses")

                    Using cmd As New HanaCommand(
                        "CALL ""DBS_BANK"".""BHV_UPDATE_PAYMENT_STATUS""()", conn)
                        cmd.ExecuteNonQuery()
                    End Using

                    Logger.Log("[" & runId & "] Bank status updated")
                End If

                '-------------------------
                ' Check completed payments
                '-------------------------
                If Not HasCompletedPayments(conn) Then
                    Logger.Log("[" & runId & "] No completed payments. Exit.")
                    Exit Sub
                End If

                '=========================
                ' SAP CONNECT
                '=========================
                oCompany = SapCompanyManager.Connect(sapSchema)
                Logger.Log("[" & runId & "] SAP Connected")

                ProcessCompletedPayments(conn, runId)

            End Using

        Catch ex As Exception
            Logger.Log("[" & runId & "] FATAL: " & ex.ToString())

        Finally
            SapCompanyManager.Disconnect(oCompany)
            Logger.Log("========== [END] RunId=" & runId & " ==========")
            Logger.EndSession()
        End Try

    End Sub

    '============================================================
    Function HasStatusUpdates(conn As HanaConnection) As Boolean
        Using cmd As New HanaCommand(
            "SELECT 1 FROM ""DBS_BANK"".""PENDING_PAYMENTS"" WHERE ""BankStatus""='PENDING' LIMIT 1", conn)
            Return cmd.ExecuteScalar() IsNot Nothing
        End Using
    End Function

    Function HasCompletedPayments(conn As HanaConnection) As Boolean
        Using cmd As New HanaCommand(
            "SELECT 1 FROM ""DBS_BANK"".""PENDING_PAYMENTS"" " &
            "WHERE ""BankStatus""='COMPLETED' AND IFNULL(""Processed"",'N')<>'Y' LIMIT 1", conn)
            Return cmd.ExecuteScalar() IsNot Nothing
        End Using
    End Function

    '============================================================
    Sub ProcessCompletedPayments(conn As HanaConnection, runId As String)

        Logger.Log("[" & runId & "] Fetching completed payments")

        Using cmd As New HanaCommand(
        "CALL ""DBS_BANK"".""BHV_GET_COMPLETED_PAYMENTS""()", conn)

            Using reader = cmd.ExecuteReader()

                While reader.Read()

                    Dim id As String = reader("Id").ToString()
                    Dim vendor As String = reader("VendorCode").ToString()
                    Dim amount As Double = Double.Parse(
                        reader("PaymentAmount").ToString().Replace(",", ""),
                        System.Globalization.CultureInfo.InvariantCulture)

                    Dim txnDate As Date = GetApprovedDate(conn, id)

                    Logger.Log($"[{runId}] Processing ID={id} | Date={txnDate:yyyy-MM-dd}")

                    Try
                        Dim paymentDocEntry As Integer = CreateOutgoingPayment(vendor, amount, txnDate)

                        UpdatePaymentProcessed(id, paymentDocEntry, conn)

                        Logger.Log($"[{runId}] SUCCESS | ID={id} | DocEntry={paymentDocEntry}")

                    Catch ex As Exception

                        UpdatePaymentError(id, ex.Message, conn)

                        Logger.Log($"[{runId}] ERROR | ID={id} | {ex.Message}")

                    End Try

                End While

            End Using
        End Using

    End Sub

    '============================================================
    Function GetApprovedDate(conn As HanaConnection, id As String) As Date

        Using cmd As New HanaCommand(
        "SELECT ""ApprovedAt"" FROM ""DBS_BANK"".""PENDING_PAYMENTS"" WHERE ""Id""=?", conn)

            cmd.Parameters.AddWithValue("p1", id)

            Dim result = cmd.ExecuteScalar()

            If result IsNot Nothing Then
                Return Convert.ToDateTime(result).Date
            End If

        End Using

        Return DateTime.Today

    End Function

    '============================================================
    Function CreateOutgoingPayment(vendor As String, amount As Double, txnDate As Date) As Integer

        Dim oRS As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
        oRS.DoQuery($"SELECT ""CardType"" FROM ""OCRD"" WHERE ""CardCode""='{vendor}'")

        Dim cardType As String = If(Not oRS.EoF, oRS.Fields.Item("CardType").Value.ToString(), "")

        If cardType = "S" OrElse cardType = "C" Then

            Dim oPayment As Payments = oCompany.GetBusinessObject(BoObjectTypes.oVendorPayments)

            oPayment.CardCode = vendor
            oPayment.DocDate = txnDate
            oPayment.DocType = If(cardType = "S", BoRcptTypes.rSupplier, BoRcptTypes.rCustomer)
            oPayment.TransferAccount = DBSBankGLAccount
            oPayment.TransferSum = amount
            oPayment.TransferDate = txnDate
            oPayment.Remarks = "DBS Payment - " & vendor

            If oPayment.Add() <> 0 Then ThrowSapError()

        Else

            Dim controlAcct As String = "210001"

            oRS.DoQuery($"SELECT ""DebPayAcct"" FROM ""OCRD"" WHERE ""CardCode""='{vendor}'")
            If Not oRS.EoF Then controlAcct = oRS.Fields.Item("DebPayAcct").Value.ToString()

            Dim oJE As JournalEntries = oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries)

            oJE.ReferenceDate = txnDate
            oJE.DueDate = txnDate
            oJE.Memo = "DBS Bank Payment - " & vendor

            oJE.Lines.AccountCode = controlAcct
            oJE.Lines.ShortName = vendor
            oJE.Lines.Debit = amount
            oJE.Lines.Add()

            oJE.Lines.AccountCode = DBSBankGLAccount
            oJE.Lines.ShortName = DBSBankGLAccount
            oJE.Lines.Credit = amount

            If oJE.Add() <> 0 Then ThrowSapError()

        End If

        Dim docEntry As String = ""
        oCompany.GetNewObjectCode(docEntry)

        Return Convert.ToInt32(docEntry)

    End Function

    '============================================================
    Sub ThrowSapError()
        Dim errCode As Integer = 0
        Dim errMsg As String = ""
        oCompany.GetLastError(errCode, errMsg)
        Throw New Exception(errCode & " - " & errMsg)
    End Sub

    '============================================================
    Sub UpdatePaymentProcessed(id As String, docEntry As Integer, conn As HanaConnection)

        Using cmd As New HanaCommand(
            "CALL ""DBS_BANK"".""BHV_MARK_PAYMENT_PROCESSED""(?,?)", conn)

            cmd.Parameters.AddWithValue("p_id", id)
            cmd.Parameters.AddWithValue("p_docnum", docEntry)
            cmd.ExecuteNonQuery()

        End Using

        ' 🔥 Clear error
        Using cmd2 As New HanaCommand(
            "UPDATE ""DBS_BANK"".""PENDING_PAYMENTS"" SET ""ErrorDescription""='' WHERE ""Id""=?", conn)

            cmd2.Parameters.AddWithValue("p_id", id)
            cmd2.ExecuteNonQuery()

        End Using

    End Sub

    Sub UpdatePaymentError(id As String, errorMsg As String, conn As HanaConnection)

        Dim safeError = If(errorMsg.Length > 500, errorMsg.Substring(0, 500), errorMsg)

        Using cmd As New HanaCommand(
            "UPDATE ""DBS_BANK"".""PENDING_PAYMENTS"" SET ""ErrorDescription""=? WHERE ""Id""=?", conn)

            cmd.Parameters.AddWithValue("p_error", safeError)
            cmd.Parameters.AddWithValue("p_id", id)
            cmd.ExecuteNonQuery()

        End Using

    End Sub
    Public Sub RunOutgoing()
        Main()
    End Sub
End Module