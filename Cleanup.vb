Imports System
Imports Sap.Data.Hana
Imports Bhavya.SAP.Core

Module Cleanup
    Sub Main()
        Dim connString As String = "Server=10.10.0.113:30015;UserID=SYSTEM;Password=B1sap#2025"
        
        Logger.StartSession()
        Logger.Log("========== [Cleanup START] " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " ==========")
        Try
            Using conn As New HanaConnection(connString)
                conn.Open()
                Logger.Log("Connected to HANA.")
                
                ' Create/Update the procedure just in case
                Dim createProcSql As String = "CREATE OR REPLACE PROCEDURE ""DBS_BANK"".""BHV_UPDATE_PROCESSED_RECONCILED"" () " &
                    "LANGUAGE SQLSCRIPT " &
                    "SQL SECURITY INVOKER " &
                    "AS " &
                    "BEGIN " &
                    "    UPDATE ""DBS_BANK"".""PENDING_PAYMENTS"" " &
                    "    SET " &
                    "        ""Processed"" = CASE WHEN ""Processed"" = 'N' THEN 'F' ELSE ""Processed"" END, " &
                    "        ""Reconciled"" = CASE WHEN ""Reconciled"" = 'N' THEN 'F' ELSE ""Reconciled"" END " &
                    "    WHERE TO_DATE(""CreatedAt"") <= ADD_DAYS(CURRENT_DATE, -2) " &
                    "      AND (""Processed"" = 'N' OR ""Reconciled"" = 'N'); " &
                    "END;"
                
                Using cmdCreate As New HanaCommand(createProcSql, conn)
                    cmdCreate.ExecuteNonQuery()
                End Using
                
                ' Execute the cleanup procedure
                Logger.Log("Running BHV_UPDATE_PROCESSED_RECONCILED...")
                Using cmdExec As New HanaCommand("CALL ""DBS_BANK"".""BHV_UPDATE_PROCESSED_RECONCILED""()", conn)
                    cmdExec.ExecuteNonQuery()
                End Using
                
                Logger.Log("Cleanup completed successfully.")
            End Using
        Catch ex As Exception
            Logger.Log("FATAL ERROR: " & ex.ToString())
        Finally
            Logger.Log("========== [Cleanup END] " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " ==========")
        End Try
        Logger.EndSession()
    End Sub
End Module
