Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

Public Class TC9_Export
    Private Shared _ErrMsg As String = String.Empty

    Public Shared Function TC9_DoExport(pConnStr As String, pOutputPath As String, pShiftDays As Integer, pSupplierID As Integer, ByRef pError As String) As Boolean

        Dim l_CN As SqlConnection _
            , l_Cmd As SqlCommand _
            , l_param As SqlParameter _
            , dr As SqlDataReader

        Try

            l_CN = New SqlConnection(pConnStr)
            l_CN.Open()
            l_Cmd = New SqlCommand("p_edi_order_export", l_CN)
            With l_Cmd
                .CommandType = CommandType.StoredProcedure
                l_param = .Parameters.Add("@supplier_id", SqlDbType.Int)
                l_param.Direction = ParameterDirection.Input
                l_param.Value = pSupplierID
                l_param = .Parameters.Add("@shift_day", SqlDbType.Int)
                l_param.Direction = ParameterDirection.Input
                l_param.Value = pShiftDays
                l_param = .Parameters.Add("@err_msg", SqlDbType.VarChar, -1)
                l_param.Direction = ParameterDirection.Output

                dr = .ExecuteReader()

            End With

            If dr.HasRows Then
                If ProcessOrders(dr, pOutputPath) Then
                    Return True
                Else
                    _ErrMsg &= vbNewLine + l_Cmd.Parameters("@err_msg").Value
                End If
            Else
                _ErrMsg &= vbNewLine & "No orders to process"
            End If
            dr.Close()
            l_CN.Close()

            Return True
        Catch ex As Exception
            _ErrMsg = ex.Message & " " & ex.Source
        Finally

            pError = _ErrMsg
        End Try
        Return False
    End Function

    Private Shared Function ProcessOrders(pDR As SqlDataReader, pOutputPath As String) As Boolean
        Dim l_Prev_Acc As String = String.Empty _
        , l_FileName As String = String.Empty _
        , writer As StreamWriter = Nothing _
        , lMHD As Integer = 1 _
        , lLineCount As Integer _
        , l_ProductCount As Integer _
        , l_Path As String = String.Empty _
        , l_Archive As String = String.Empty
        Try

            While pDR.Read
                If String.IsNullOrEmpty(l_Path) Then
                    l_Path = pOutputPath
                    If Not String.IsNullOrEmpty(pDR("supp_acc")) Then
                        l_Path &= "\" & pDR("supp_acc")
                    End If
                    l_Archive = l_Path & "\Archive"

                    If Directory.Exists(l_Path) Then
                    Else
                        Directory.CreateDirectory(l_Path)
                    End If

                    If Directory.Exists(l_Archive) Then
                    Else
                        Directory.CreateDirectory(l_Archive)
                    End If
                End If

                If Not String.Equals(pDR("acc_num"), l_Prev_Acc, StringComparison.CurrentCultureIgnoreCase) Then
                    'control break
                    If Not writer Is Nothing Then
                        With New StringBuilder
                            .Append("OTR=" + l_ProductCount.ToString) : .AppendLine("'")
                            .Append("MTR=" + lLineCount.ToString) : .AppendLine("'")
                            lLineCount = 0
                            .Append("MHD=") : .Append(lMHD.ToString) : .AppendLine("+ORDTLR:9'") : lLineCount += 1
                            .AppendLine("OFT=1'") : lLineCount += 1
                            .Append("MTR=") : .Append(lLineCount.ToString) : .AppendLine("'")
                            .Append("END=") : .Append(lMHD.ToString) : .Append("'")
                            writer.Write(.ToString())
                        End With
                        writer.Close()

                        File.Copy(l_Path & "\" & l_FileName, l_Archive & "\" & l_FileName, True)
                    End If

                    l_FileName = pDR("delivery_date").ToString & "_" & pDR("acc_num").ToString.Trim & ".EDI"

                    If File.Exists(l_Path & "\" & l_FileName) Then
                        File.Delete(l_Path & "\" & l_FileName)
                    End If
                    writer = File.CreateText(l_Path & "\" & l_FileName)
                    writer.AutoFlush = True

                    l_Prev_Acc = pDR("acc_num")
                    lLineCount = 0 : lMHD = 1 : l_ProductCount = 0

                    ' Write heading
                    With New StringBuilder
                        .Append("STX=ANA:1+5000000000000:WEST COUNTRY MILK+") 'standard WCM header
                        lLineCount += 1
                        Select Case pDR("supplier")
                            Case "Medina Dairy"
                                .Append("5026091000007:MEDINA FOOD SERVICE+") 'customer specific part
                            Case "Payne's Dairies"
                                .Append("5060043302795:PAYNES DAIRIES+") 'customer specific part
                        End Select
                        'pOrder.
                        .Append(Now.ToString("yyMMdd")) : .Append(":") : .Append(Now.ToString("hhmmss")) : .Append("+") 'current date and time
                        .Append("") 'Sender Transmission Reference
                        .Append("") : .AppendLine("++ORDHDR'") : lLineCount += 1 'Recipient's Transmission Reference
                        .Append("MHD=") : .Append(lMHD.ToString) : .AppendLine("+ORDHDR:9'") : lLineCount += 1
                        lMHD += 1

                        .AppendLine("TYP=0430+NEW-ORDERS'") : lLineCount += 1
                        Select Case pDR("supplier")
                            Case "Medina Dairy"
                                .AppendLine("SDT=5026091000007:MD3+MEDINA FOOD SERVICE+A8-A11 New Covent Garden Market::London::SW8 5EE'") 'customer specific part
                            Case "Payne's Dairies"
                                .AppendLine("SDT=5060043302795:CP1+PAYNES DAIRIES+Bar Lane:Boroughbridge:North Yorkshire::YO51 9LU'") 'customer specific part
                        End Select
                        lLineCount += 1

                        .AppendLine("CDT=5000000000000+West Country Milk Consortium+OTTER BUILDING:GRENADIER ROAD:EXETER BUSINESS PARK:EXETER:EX1 3QN'") : lLineCount += 1
                        .Append("FIL=") : .Append(pDR("order_id").ToString) : .Append("+1+") : .Append(Now.ToString("yyMMdd")) : .AppendLine("'") : lLineCount += 1

                        .Append("MTR=") : .Append(lLineCount.ToString) : .AppendLine("'")

                        lLineCount = 0 'reset line count before next header
                        .Append("MHD=") : .Append(lMHD.ToString) : .AppendLine("+ORDERS:9'") : lLineCount += 1
                        lMHD += 1
                        .Append("CLO=:") : .Append(pDR("acc_num").ToString.Trim) : .Append("+") : .Append(pDR("customer").ToString) : .AppendLine("'") : lLineCount += 1 'might need to add in the customer address 
                        .Append("ORD=") : .Append(pDR("order_num").ToString) : .Append("::") : .Append(Now.ToString("yyMMdd")) : .AppendLine("'") : lLineCount += 1

                        .Append("DIN=") : .Append(pDR("delivery_date").ToString) : .AppendLine("'") : lLineCount += 1

                        writer.Write(.ToString)
                    End With
                End If

                'write details
                l_ProductCount += 1
                With New StringBuilder
                    .Append("OLD=") : .Append(l_ProductCount.ToString) : .Append("+++:") : .Append(pDR("pr_code").ToString.Trim)

                    'bMappingOk = mdctProductMapping.ContainsKey(sCode_Supplier & lInvoice.CostBandID)

                    'If bMappingOk Then
                    '    If lInvoice.SupplierId = 84388339 Then 'FR6 
                    '        structMap = mdctProductMapping_FR6.Item(sCode_Supplier & lInvoice.CostBandID)
                    '    Else
                    '        structMap = mdctProductMapping.Item(sCode_Supplier & lInvoice.CostBandID)
                    '    End If
                    'End If

                    '.Append("MEDINA CODE")    'need to get the medina product code

                    .Append("++1+") : .Append(pDR("Qty")) : .Append("+++") 'unit cost, special price,to follow indicator
                    .Append(pDR("product")) : .AppendLine("'") : lLineCount += 1 'product description

                    writer.Write(.ToString)
                End With
            End While
            '
            If Not writer Is Nothing Then
                With New StringBuilder
                    .Append("OTR=" + l_ProductCount.ToString) : .AppendLine("'")
                    .Append("MTR=" + lLineCount.ToString) : .AppendLine("'")
                    lLineCount = 0
                    .Append("MHD=") : .Append(lMHD.ToString) : .AppendLine("+ORDTLR:9'") : lLineCount += 1
                    .AppendLine("OFT=1'") : lLineCount += 1
                    .Append("MTR=") : .Append(lLineCount.ToString) : .AppendLine("'")
                    .Append("END=") : .Append(lMHD.ToString) : .Append("'")
                    writer.Write(.ToString())
                End With
                writer.Close()

                File.Copy(l_Path & "\" & l_FileName, l_Archive & "\" & l_FileName, True)
            End If

            Return True
        Catch ex As Exception
            _ErrMsg = ex.Message & " " & ex.Source
        Finally
            If Not writer Is Nothing Then
                writer.Close()
            End If
            pDR.Close()
        End Try
        Return False
    End Function

End Class
