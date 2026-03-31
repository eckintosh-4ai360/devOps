'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>

Dim strPatName, strInsrNo, dtBirth, strTelNo

printlayoutName = "ItemPurOrderDashboard"
tittle = "General Store Purchasing Dashboard"
MainTable = ""

sup = Trim(Request.queryString("supID"))
mth = Trim(Request.queryString("wrkmID"))

response.Clear

response.write "<!DOCTYPE html>"
response.write "<html>"
response.write "  <head>"
response.write "     <title>" & tittle & "</title>"
response.write "    <link rel='stylesheet' type='text/css' href='fontawesome-free-6.4.0-web/css/all.css'/>"
StylesAdded
response.write "  </head>"
response.write "  <body>"

If mth = "" And sup = "" Then
    mth = GetSupWrkmnt()
    sqlWhere = "WHERE WorkingMonthID = '" & GetSupWrkmnt() & "' "
Else
    sqlWhere = "WHERE "
End If

cnt = 0

If Len(sup) > 0 And sup <> "All" Then
    cnt = cnt + 1
    If cnt > 1 Then
      sqlWhere = sqlWhere & " "
    End If
    sqlWhere = sqlWhere & " SupplierID = '" & sup & "'"
End If
If Len(mth) > 0 And mth <> GetSupWrkmnt() And mth <> "All" Then
    cnt = cnt + 1
    If cnt > 1 Then
      sqlWhere = sqlWhere & " and "
    End If
    sqlWhere = sqlWhere & " WorkingMonthID ='" & mth & "'"
ElseIf mth = "All" Then
    sqlWhere = sqlWhere
End If

If mth = "All" And sup = "" Then
    sqlWhere = ""
End If

sql = ""
sql = sql & " SELECT ItemPurOrderID, SupplierTypeID, PurchaseOrderDate, SupplierID, ItemPurOrderTypeID, ItemPurOrderStatusID, ItemCategoryID, SystemuserID, TransProcessValID, PurchaseOrderInfo1"
sql = sql & " FROM ItemPurOrder " & sqlWhere & " order by PurchaseOrderDate desc  "

DisplayReport sql

ScpAdded
TableEXScp
'response.write "    <script src='https://kit.fontawesome.com/91bd23ceba.js'></script>"
response.write "  </body>"
response.write "</html>"

Sub DisplayReport(querry)

    Dim rst, sql, rst2, sql2, sqlWhere
    Dim iCnt, recStyle, href, data

    Set rst = CreateObject("ADODB.Recordset")

    With rst
        rst.open qryPro.FltQry(querry), conn, 3, 4
        iCnt = 1
        Dim patID

        If rst.RecordCount > 0 Then
            rst.MoveFirst

            TopHead tittle
            response.write "      <div class='cnt-body p15 m25 pt0 mt0'>"
            TableEX
            response.write "        <div class='tab-table tbl-main mil'>"
            response.write "          <div class='table-holder'>"
            response.write "            <table class='tbl'>"
            response.write "              <thead class='tbl-h'>"
            response.write "                <tr>"
            response.write "                  <th data-sort='none' class='fh t-left'>No.</th>"
            response.write "                  <th data-sort='none' class='fh t-left'>Purchase No.</th>"
            response.write "                  <th data-sort='none' class='fh t-left'>Supplier</th>"
            response.write "                  <th data-sort='none' class='fh t-left'>Details</th>"
            response.write "                  <th data-sort='none' class='fh t-left'>Approval Status</th>"
            response.write "                  <th data-sort='none' class='fh t-left'>Acceptance</th>"
            response.write "              </thead>"
            response.write "              <tbody class='tbl-b'>"
            Do While Not rst.EOF
                purID = rst.fields("ItemPurOrderID")
                supp = rst.fields("SupplierID")
                suppt = rst.fields("SupplierTypeID")
                typ = rst.fields("ItemPurOrderTypeID")
                pDt = rst.fields("PurchaseOrderDate")
                purfr = rst.fields("ItemPurOrderStatusID")
                transproStage = rst.fields("TransProcessValID")
                cat = rst.fields("ItemCategoryID")
                usr = rst.fields("SystemuserID")
                flow = rst.fields("PurchaseOrderInfo1")
                pos = rst.AbsolutePosition

                href = "wpgItemPurOrder.asp?PageMode=ProcessSelect&ItemPurOrderID=" & purID
                phref = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=ItemPurOrder&PositionForTableName=ItemPurOrder&ItemPurOrderID=" & purID

                response.write "                <tr>"
                response.write "                  <td>" & CStr(iCnt) & "</td>"
                response.write "                  <td><div class='df fc fs'>"
                response.write "                    <div class='bold p5'>" & Addlnk(purID, href, "pm aln") & "</div>"
                response.write "                    <div class='p5'>" & Addlnk(FormatDateDetail(pDt), href, "bnm") & "</div>"
If Approved(purID, "ItemPurOrderPro-T005") Then
                response.write "                    <div class='p5 bwrn'>*Record Edited</div>"
End If
                response.write "                  </div></td>"
                response.write "                  <td><div class='df fc fs'>"
                response.write "                    <div class='p5 bold'>" & Addlnk(GetComboName("Supplier", supp), href, "bnm") & "</div>"
                response.write "                    <div class='p5'>" & Addlnk(GetComboName("SupplierType", suppt), href, "bnm") & "</div>"
                response.write "                  </div></td>"
                response.write "                  <td><div class='df fc fs'>"
                response.write "                    <div class='p5'><span class='pr5 bold'>Type:</span>" & GetComboName("ItemPurOrderType", typ) & "</div>"
                response.write "                    <div class='p5'><span class='pr5 bold'>Item Category:</span>" & GetComboName("ItemCategory", cat) & "</div>"
                response.write "                    <div class='p5 df rlt'><span class='pr5 bold wn'>Order Items:</span>"
                    response.write "                <div class='btsty bs showsrv" & CStr(iCnt) & "' style='margin: 5px;'>View Items</div>"
                    'If UCase(transproStage) <> UCase("ItemPurOrderPro-T001") And UCase(transproStage) <> UCase("ItemPurOrderPro-T005") Then
                        response.write "                <div>" & Addlnk("Print", phref, "bs") & "</div>"
                    'End If
                    response.write "                <div class='df p5 abs licnt dn srvlist" & CStr(iCnt) & "'>"
                    response.write "                  <div class='df fc rlt alf-end'>"
                    response.write "                    <div class='df j-sb w100'><div class='t2 f14 plr5'>Order Items</div><div class='clsbtt clsbtn" & CStr(iCnt) & "'><div class='btsty bwrn'>close</div></div></div>"
                    response.write "                    <table class='tbl-Sub' style='border-collapse: collapse; width: max-content;'>"
                    response.write "                      <thead>"
                    response.write "                      <tr>"
                    response.write "                        <th class='fh1'>No</th>"
                    response.write "                        <th class='fh1 t-left'>Item / Description</th>"
                    response.write "                        <th class='fh1 t-right'>Order Qty</th>"
                    response.write "                        <th class='fh1 t-right'>Purch. Price</th>"
                    response.write "                        <th class='fh1 t-right'>Order Amt.</th>"
                    response.write "                        <th class='fh1 t-left'>UOM</th>"
                    response.write "                      </tr>"
                    response.write "                    </thead>"
                    DisplayRequestedItems purID
                    response.write "                    </table>"
                    response.write "                 </div>"
                    response.write "                </div></div>"
                response.write "                  </div></td>"
                response.write "                  <td><div class='df fc fs'>"
                response.write "                    <div class='p5'><span class='pr5 bold'>Requested by:</span>" & GetComboName("Staff", GetComboNameFld("SystemUser", usr, "StaffID")) & "</div>"
            If UCase(flow) = "P002" Or UCase(flow) = "" Then 'Hence Workflow
                If RequiresCEOApproval(purID) Then
                    If Approved(purID, "ItemPurOrderPro-T011") Then
                        response.write "                    <div class='p5'><span class='pr5 bold'>Approved by:</span>" & getUserAppv(purID, "ItemPurOrderPro-T011") & "</div>"
                    Else
                        response.write "                    <div class='df w100'><div class='p5 f1'><span class='datastat pnd'>Pending Inventory Supervisor Approval</span></div><div class='f1 df f-end'>"
                        sTb2 = "ItemPurOrderPro"
                        If HasAccessRight(uName, "frm" & sTb2, "New") And Glob_HasTransProcessAccess(sTb2, uName, "T001", "T011") Then
                            lnkText = "Approve Purchase Request"
                            lnkUrl = "wpgItemPurOrderPro.asp?PageMode=AddNew&TransProcessVal2ID=ItemPurOrderPro-T011&PullupData=ItemPurOrderID||" & purID
                            response.write Addlnk(lnkText, lnkUrl, "bss")
                        End If
                        response.write "                  </div></div>"
                    End If
                Else
                    If Approved(purID, "ItemPurOrderPro-T002") Then
                        response.write "                    <div class='p5'><span class='pr5 bold'>Approved by:</span>" & getUserAppv(purID, "ItemPurOrderPro-T002") & "</div>"
                    Else
                        response.write "                    <div class='df w100'><div class='p5 f1'><span class='datastat pnd'>Pending Inventory Supervisor Approval</span></div><div class='f1 df f-end'>"
                        sTb2 = "ItemPurOrderPro"
                        If HasAccessRight(uName, "frm" & sTb2, "New") And Glob_HasTransProcessAccess(sTb2, uName, "T001", "T002") Then
                            lnkText = "Approve Purchase Request"
                            lnkUrl = "wpgItemPurOrderPro.asp?PageMode=AddNew&TransProcessVal2ID=ItemPurOrderPro-T002&PullupData=ItemPurOrderID||" & purID
                            response.write Addlnk(lnkText, lnkUrl, "bss")
                        End If
                        response.write "                  </div></div>"
                    End If
                End If

                If Not RequiresCEOApproval(purID) Then
                    If Approved(purID, "ItemPurOrderPro-T003") Then
                        response.write "                    <div class='p5'><span class='pr5 bold'>Approved by:</span>" & getUserAppv(purID, "ItemPurOrderPro-T003") & "</div>"
                    Else
                        response.write "                    <div class='df w100'><div class='p5 f1'><span class='datastat pnd'>Pending Administrator Approval</span></div><div class='f1 df f-end'>"
                        sTb2 = "ItemPurOrderPro"
                        If HasAccessRight(uName, "frm" & sTb2, "New") And Glob_HasTransProcessAccess(sTb2, uName, "T002", "T003") And Approved(purID, "ItemPurOrderPro-T002") Then
                            lnkText = "Approve Purchase Request"
                            lnkUrl = "wpgItemPurOrderPro.asp?PageMode=AddNew&TransProcessVal2ID=ItemPurOrderPro-T003&PullupData=ItemPurOrderID||" & purID
                            response.write Addlnk(lnkText, lnkUrl, "bss")
                        End If
                        response.write "                  </div></div>"
                    End If

                End If

                If RequiresCEOApproval(purID) Then
                    If Approved(purID, "ItemPurOrderPro-T004") Then
                        response.write "                    <div class='p5'><span class='pr5 bold'>Approved by:</span>" & getUserAppv(purID, "ItemPurOrderPro-T004") & "</div>"
                    Else
                        'If UCase(flow) = "P002" Then
                            response.write "                    <div class='df w100'><div class='p5 f1'><span class='datastat pnd'>Pending CEO/CFO Approval</span></div><div class='f1 df f-end'>"
                            sTb2 = "ItemPurOrderPro"
                            If HasAccessRight(uName, "frm" & sTb2, "New") And Glob_HasTransProcessAccess(sTb2, uName, "T011", "T004") Then
                                lnkText = "Approve Purchase Request"
                                lnkUrl = "wpgItemPurOrderPro.asp?PageMode=AddNew&TransProcessVal2ID=ItemPurOrderPro-T004&PullupData=ItemPurOrderID||" & purID
                                response.write Addlnk(lnkText, lnkUrl, "bss")
                            End If
                            response.write "                  </div></div>"
                        'End If
                    End If
                End If

                If (UCase(transproStage) = "ITEMPURORDERPRO-T005" And Approved(purID, "ItemPurOrderPro-T004")) Or (UCase(transproStage) = "ITEMPURORDERPRO-T005" And Approved(purID, "ItemPurOrderPro-T003")) Or (UCase(transproStage) = "ITEMPURORDERPRO-T005" And Approved(purID, "ItemPurOrderPro-T002")) Then
                    If Approved(purID, "ItemPurOrderPro-T007") Then
                        response.write "                    <div class='p5'><span class='pr5 bold'>Approved by:</span>" & getUserAppv(purID, "ItemPurOrderPro-T007") & "</div>"
                    Else
                        response.write "                    <div class='df w100'><div class='p5 f1'><span class='datastat pnd'>Pending Approval After Change Request</span></div><div class='f1 df f-end'>"
                        sTb2 = "ItemPurOrderPro"
                        If HasAccessRight(uName, "frm" & sTb2, "New") And Glob_HasTransProcessAccess(sTb2, uName, "T005", "T007") Then
                            lnkText = "Approve Change Request"
                            lnkUrl = "wpgItemPurOrderPro.asp?PageMode=AddNew&TransProcessVal2ID=ItemPurOrderPro-T007&PullupData=ItemPurOrderID||" & purID
                            response.write Addlnk(lnkText, lnkUrl, "bss")
                        End If
                        response.write "                  </div></div>"
                    End If
                End If
            Else
                If Approved(purID, "ItemPurOrderPro-T008") Then
                    response.write "                    <div class='p5'><span class='pr5 bold'>Approved by:</span>" & getUserAppv(purID, "ItemPurOrderPro-T008") & "</div>"
                Else
                    response.write "                    <div class='df w100'><div class='p5 f1'><span class='datastat pnd'>Pending Inventory Head Approval</span></div><div class='f1 df f-end'>"
                    sTb2 = "ItemPurOrderPro"
                    If HasAccessRight(uName, "frm" & sTb2, "New") And Glob_HasTransProcessAccess(sTb2, uName, "T001", "T008") Then
                        lnkText = "Approve Purchase Request"
                        lnkUrl = "wpgItemPurOrderPro.asp?PageMode=AddNew&TransProcessVal2ID=ItemPurOrderPro-T008&PullupData=ItemPurOrderID||" & purID
                        response.write Addlnk(lnkText, lnkUrl, "bss")
                    End If
                    response.write "                  </div></div>"
                End If

                If (UCase(transproStage) = "ITEMPURORDERPRO-T005" And Approved(purID, "ItemPurOrderPro-T008")) Then
                    If Approved(purID, "ItemPurOrderPro-T007") Then
                        response.write "                    <div class='p5'><span class='pr5 bold'>Approved by:</span>" & getUserAppv(purID, "ItemPurOrderPro-T007") & "</div>"
                    Else
                        response.write "                    <div class='df w100'><div class='p5 f1'><span class='datastat pnd'>Pending Approval After Change Request</span></div><div class='f1 df f-end'>"
                        sTb2 = "ItemPurOrderPro"
                        If HasAccessRight(uName, "frm" & sTb2, "New") And Glob_HasTransProcessAccess(sTb2, uName, "T005", "T007") Then
                            lnkText = "Approve Purchase Request"
                            lnkUrl = "wpgItemPurOrderPro.asp?PageMode=AddNew&TransProcessVal2ID=ItemPurOrderPro-T007&PullupData=ItemPurOrderID||" & purID
                            response.write Addlnk(lnkText, lnkUrl, "bss")
                        End If
                        response.write "                  </div></div>"
                    End If
                End If
            End If
                If UCase(transproStage) > "ITEMPURORDERPRO-T001" And UCase(transproStage) <> "ITEMPURORDERPRO-T005" Then
                    sTb2 = "ItemPurOrderPro"
                    If HasAccessRight(uName, "frm" & sTb2, "New") And (Glob_HasTransProcessAccess(sTb2, uName, "T002", "T005") Or Glob_HasTransProcessAccess(sTb2, uName, "T003", "T005") Or Glob_HasTransProcessAccess(sTb2, uName, "T011", "T005") Or Glob_HasTransProcessAccess(sTb2, uName, "T004", "T005") Or Glob_HasTransProcessAccess(sTb2, uName, "T008", "T005")) Then
                        lnkText = "Return to Requester"
                        lnkUrl = "wpgItemPurOrderPro.asp?PageMode=AddNew&TransProcessVal2ID=ItemPurOrderPro-T005&PullupData=ItemPurOrderID||" & purID
                        response.write Addlnk(lnkText, lnkUrl, "bwrn")
                    End If
                End If

                response.write "                  </div></td>"
                response.write "                  <td>"
                DisplayAcceptedItems purID, pos
                response.write "                  </td>"
                response.write "                </tr>"

                iCnt = iCnt + 1
                response.flush
                rst.MoveNext
            Loop
            response.write "              </tbody>"
            response.write "            </table>"
            response.write "          </div>"
            response.write "        </div>"
            response.write "      </div>"
        Else

            TopHead tittle
            response.write "  <div class='t-case rlt t-title'>"
            response.write "    <div class='cnt'>"
            response.write "      <div class='t-title t3 df jc p10 m15 mlr25 pb0 pt0 dnp'>"
            response.write "        No Record Found"
            response.write "      </div>"
            response.write "    </div>"
            response.write "  </div>"
        End If

        rst.Close
    End With

    Set rst = Nothing

End Sub

Sub DisplayAcceptedItems(reqID, pos)
    Dim rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select da.*, d.ItemName from IncomingStockItems da, Items d where d.ItemID=da.ItemID and da.ItemPurOrderID='" & reqID & "' "
    sql = sql & " order by d.ItemName "

    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
            ref = "wpgIncomingStock.asp?PageMode=ProcessSelect&PullupData=ItemPurOrderID||" & reqID
            rst.MoveFirst
            response.write "                  <div class='df fc fs rls'>"
            response.write "                    <div class='p5'><span class='pr5 bold'>Accepted By:</span>" & Addlnk(GetComboName("Staff", GetComboNameFld("SystemUser", rst.fields("SystemUserID"), "StaffID")), ref, "bnm") & " <span class='st2s'>" & FormatDate(rst.fields("Entrydate")) & "</span></div>"
            response.write "                    <div class='p5 df rlt'><span class='pr5 bold'>Accepted Items:</span>"
                response.write "                <div class='btsty bs bdshow" & pos & "' style='margin: 5px;'>View Items</div>"
                'response.write "                <div>" & Addlnk("Print SRV", ihref, "bs") & "</div>"
                response.write "                <div class='df p5 abs licnt rt0 dn bdlist" & pos & "'>"
                response.write "                  <div class='df fc rlt alf-end'>"
                response.write "                    <div class='df j-sb w100'><div class='t2 f14 plr5'>Accepted Items</div><div class='clsbtt bdbtn" & pos & "'><div class='btsty bwrn'>close</div></div></div>"
                response.write "                    <table class='tbl-Sub' style='border-collapse: collapse; width: max-content;'>"
                response.write "                      <thead>"
                response.write "                      <tr>"
                response.write "                        <th class='fh1'>No</th>"
                response.write "                        <th class='fh1 t-left'>Item / Description</th>"
                response.write "                        <th class='fh1 t-right'>Order Qty</th>"
                response.write "                        <th class='fh1 t-right'>supply Quantity</th>"
                response.write "                        <th class='fh1 t-right'>Purchase Price</th>"
                response.write "                        <th class='fh1 t-left'>Total Cost</th>"
                response.write "                      </tr>"
                response.write "                    </thead>"

            Do While Not rst.EOF
                response.write "<tr>"
                response.write "<td>" & rst.AbsolutePosition & "</td>"
                response.write "<td>" & rst.fields("ItemName") & "</td>"
                response.write "<td class='fh t-right'>" & FormatNumber(rst.fields("Qty"), 1) & "</td>"
                response.write "<td class='fh t-right'>" & FormatNumber(rst.fields("ReturnQty"), 1) & "</td>"
                response.write "<td class='fh t-right'>" & FormatNumber(rst.fields("BulkUnitCost"), 1) & "</td>"
                response.write "<td class='fh t-right'>" & FormatNumber(rst.fields("TotalCost"), 1) & "</td>"
                response.write "</tr>"
                response.flush
                rst.MoveNext
            Loop
                response.write "                    </table>"
                response.write "                 </div>"
                response.write "                </div></div>"
                'If Approved(reqID, "ItemPurOrderPro-T002") Then 'by Mike
                    If SupplyNotComplete(reqID) Then
                        sTb2 = "IncomingStock"
                        If HasAccessRight(uName, "frm" & sTb2, "New") Then
                            lnkText = "Accept Remaining Items"
                            lnkUrl = "wpg" & sTb2 & ".asp?PageMode=AddNew&PullupData=ItemPurOrderID||" & reqID
                            response.write Addlnk(lnkText, lnkUrl, "bss")
                        End If
                    'End If
                End If
            response.write "                  </div>"
        Else
            response.write "                  <div class='df fc fs rls'>"
            response.write "                    <div class='p5'><span class='pr5 bold'>No Accepted Items</span></div>"
            sTb2 = "IncomingStock"
            If (Approved(reqID, "ItemPurOrderPro-T002") Or Approved(reqID, "ItemPurOrderPro-T003") Or Approved(reqID, "ItemPurOrderPro-T004") Or Approved(reqID, "ItemPurOrderPro-T008")) Then 'by Mike
                 If HasAccessRight(uName, "frm" & sTb2, "New") Then
                    lnkText = "Accept Items to My Store"
                    lnkUrl = "wpg" & sTb2 & ".asp?PageMode=AddNew&PullupData=ItemPurOrderID||" & reqID
                    response.write Addlnk(lnkText, lnkUrl, "bss")
                 End If
            End If
            response.write "                  </div>"
        End If
        rst.Close
    End With
    Set rst = Nothing
End Sub

Sub DisplayRequestedItems(reqID)
    Dim rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    sql = "select dr.*, d.ItemName, u.UnitOfMeasureName from ItemPurOrderItems dr, Items d, UnitOfMeasure u Where d.ItemID=dr.ItemID And u.UnitOfMeasureID=d.UnitOfMeasureID "
    sql = sql & " And dr.ItemPurOrderID='" & reqID & "' order by d.ItemName "
    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
            rst.MoveFirst

            Do While Not rst.EOF
                response.write "<tr>"
                response.write "<td>" & rst.AbsolutePosition & "</td>"
                response.write "<td>" & rst.fields("ItemName") & "</td>"
                response.write "<td class='fh t-right'>" & FormatNumber(rst.fields("OrderQuantity"), 1) & "</td>"
                response.write "<td class='fh t-right'>" & FormatNumber(rst.fields("OrderAmount1"), 1) & "</td>"
                response.write "<td class='fh t-right'>" & FormatNumber(rst.fields("OrderAmount2"), 1) & "</td>"
                response.write "<td>" & rst.fields("UnitOfMeasureName") & "</td>"
                response.write "</tr>"
                response.flush
                rst.MoveNext
            Loop
        Else

        End If
        rst.Close
    End With
    Set rst = Nothing
End Sub

Function SupplyNotComplete(reqID)
    Dim rst, sql, ot, otn
    otn = False
    Set rst = CreateObject("ADODB.Recordset")
    With rst
        sql = "Select Sum(ReturnQty) tot from IncomingStockItems "
        sql = sql & " WHERE ItemPurOrderID='" & reqID & "' "
        .maxrecords = 1
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            ot = rst.fields("tot")
        End If
        If ot < getReqQuantity(reqID) Then
            otn = True
        End If
        .Close
    End With
    SupplyNotComplete = otn
    Set rst = Nothing
End Function

Function RequiresCEOApproval(purID)
    Dim sql, rst, rst2, maxPurLm

    ot = False
    maxPurLm = 5000 'max purchase limit

    sql = "select * from ItemPurOrder where ItemPurOrderID='" & purID & "' "
    Set rst = CreateObject("ADODB.RecordSet")
    Set rst2 = CreateObject("ADODB.RecordSet")
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst

        sql = "select sum(OrderAmount2) as [total] "
        sql = sql & " from ( "
        sql = sql & "     select OrderAmount2 from ItemPurOrderItems where ItemPurOrderID='" & rst.fields("ItemPurOrderID") & "' "
        sql = sql & "     union all select OrderAmount2 from ItemPurOrderItems2 where ItemPurOrderID='" & rst.fields("ItemPurOrderID") & "' "
        sql = sql & " ) as Purchases"
        sql = sql & " "

        rst2.open qryPro.FltQry(sql), conn, 3, 4
        If rst2.RecordCount > 0 Then
            If Not IsNull(rst2.fields("total")) Then
                If rst2.fields("total") >= maxPurLm Then 'CEO's approval required
                    ot = True
                End If
            End If
        End If
    End If

    RequiresCEOApproval = ot
End Function

Function getReqQuantity(reqID)
    Dim rst, sql, ot
    ot = False
    Set rst = CreateObject("ADODB.Recordset")
    With rst
        sql = " Select Sum(OrderQuantity) as tot from ItemPurOrderItems "
        sql = sql & " WHERE ItemPurOrderID='" & reqID & "' "
        .maxrecords = 1
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            ot = rst.fields("tot")
        End If
        .Close
    End With
    getReqQuantity = ot
    Set rst = Nothing
End Function

Function HasApprvAcc(jSchd, transproStage)
    Dim rst, sql, ot
    ot = False
    Set rst = CreateObject("ADODB.Recordset")
    With rst
        sql = "Select top 1 TransProcessorAcc2ID from Transprocessor "
        sql = sql & " where JobscheduleID='" & jSchd & "' And TransProcessVal2ID='" & transproStage & "'"
        .maxrecords = 1
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            ot = True
        End If
        .Close
    End With
    HasApprvAcc = ot
    Set rst = Nothing
End Function

Function Approved(purID, transproStage)
    Dim rst, sql, ot
    ot = False
    Set rst = CreateObject("ADODB.Recordset")
    With rst
        sql = "Select top 1 ItemPurOrderProID from ItemPurOrderPro "
        sql = sql & " where ItemPurOrderID='" & purID & "' And TransProcessVal2ID='" & transproStage & "'"
        .maxrecords = 1
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            ot = True
        End If
        .Close
    End With
    Approved = ot
    Set rst = Nothing
End Function

Function getUserAppv(purID, transproStage)
    Dim rst, sql, ot
    ot = False
    Set rst = CreateObject("ADODB.Recordset")
    With rst
        sql = "Select top 1 SystemuserID from ItemPurOrderPro "
        sql = sql & " where ItemPurOrderID='" & purID & "' And TransProcessVal2ID='" & transproStage & "'"
        .maxrecords = 1
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            ot = GetComboName("Staff", GetComboNameFld("SystemUser", rst.fields("SystemuserID"), "StaffID"))
        End If
        .Close
    End With
    getUserAppv = ot
    Set rst = Nothing
End Function

Function GetSupWrkmnt()
    Dim rst, sql, ot
    ot = False
    Set rst = CreateObject("ADODB.Recordset")
    With rst
        sql = "Select top 1 WorkingMonthID from ItemPurOrder "
        sql = sql & " Order by WorkingMonthID Desc "
        .maxrecords = 1
        .open qryPro.FltQry(sql), conn, 3, 4
        If .RecordCount > 0 Then
            ot = rst.fields("WorkingMonthID")
        End If
        .Close
    End With
    GetSupWrkmnt = ot
    Set rst = Nothing
End Function

Function GetInact(lnkText)
    Dim html
    html = ""

    html = html & "     <div class='btsty' style='background-color: #eeeeee;'>" & lnkText & "</div>"
    Addlnk = html
End Function

Function Addlnk(lnkText, url, extraClass)
    Dim html
    html = ""

    If (extraClass = "") Or (extraClass = "bnm") Then
        html = html & "     <div class='btsty aln bnm' onclick=""openPopup('" & url & "', 800, 600)"">" & lnkText & "</div>"
    Else
        html = html & "     <div class='btsty " & extraClass & "' onclick=""openPopup('" & url & "', 800, 600)"">" & lnkText & "</div>"
    End If
    Addlnk = html
End Function

Sub TopHead(tittle)

    response.write "  <div class='t-case rlt t-title'>"
    response.write "    <div class='abs min-top df jc cp'>"
    response.write "      <div class='rlt'><div class='abs t-case-i'>Minimize</div></div>"
    response.write "      <i class='fa-solid fa-angle-up icnn'></i>"
    response.write "    </div>"
    response.write "    <div class='cnt'>"
    response.write "      <div class='t-title df fs p10 m15 mlr25 pb0 pt0 dnp'>"
    response.write "        <div class='t1' style='padding: 10px 0px 10px 55px;'>"
    response.write "          " & tittle
    response.write "        </div>"
    response.write "      </div>"
    response.write "      <div class='t-title df fs p10 m15 mlr25 pb0 pt0 dnp'>"
    response.write DetailSelectorFilter("SupplierID", "Supplier", "Supplier", "supID")
    response.write DetailSelectorFilter("WorkingMonthID", "Month", "WorkingMonth", "wrkmID")
    response.write "    <div class='df fs'>"
    response.write SubmitButton("Submit", "supID||wrkmID")
    response.write "    </div>"
    response.write "      </div>"
    ' response.write "    </div>"
    response.write "    </div>"
    response.write "  </div>"

End Sub

Function SubmitButton(title, submitIDs)
    html = " "

    html = html & " <div class='btsty bs' id='submitButton'>" & title & "</div>" & vbCrLf
    html = html & " <script>" & vbCrLf

    Dim idArray
    idArray = Split(submitIDs, "||")
    For Each ID In idArray
        html = html & "    let " & ID & " = '" & ID & "';" & vbCrLf
    Next
    html = html & "    document.getElementById('submitButton').addEventListener('click', function() {" & vbCrLf
    html = html & "        submitDetails();" & vbCrLf
    html = html & "    });" & vbCrLf
    html = html & "    function submitDetails() {" & vbCrLf
    For Each ID In idArray
        html = html & "        const " & ID & "Value = document.getElementById('" & ID & "').value;" & vbCrLf
    Next
    html = html & "        let currentURL = 'wpgPrtPrintLayoutAll.asp?PrintLayoutName=" & printlayoutName & "&PositionForTableName=WorkingDay';" & vbCrLf
    html = html & "        let newURL = currentURL + '&WorkingDayID=DAY20160401';" & vbCrLf
    For Each ID In idArray
        html = html & "        newURL += '&" & ID & "=' + " & ID & "Value;" & vbCrLf
    Next
    html = html & "        window.location.href = newURL;" & vbCrLf
    html = html & "    }" & vbCrLf
    html = html & " </script>" & vbCrLf

    SubmitButton = html
End Function

Function DetailSelector(itmName, inputtype, ID, placeHolder, extraClass)
    html = " "

    html = html & " <div class='period-selector df f1 plr25 mlr25'> "
    html = html & "     <span class='plr10 t2 mw-115'>" & itmName & "</span>"
    html = html & "     <input class='select-value1 " & extraClass & "' type='" & inputtype & "' id='" & ID & "' placeholder='" & placeHolder & "' />"
    html = html & " </div>"

    DetailSelector = html
End Function

Function DateSelector(itmName, ID, extraClass)
    html = " "

    html = html & "     <span class='plr10 t2 " & extraClass & "'>" & itmName & "</span>"
    html = html & "     <input class='select-value1' type='date' id='" & ID & "'/>"

    DateSelector = html
End Function

Function DetailSelectorFilter(itmID, itmName, table, ID)
    Dim rst, sql
    Set rst = CreateObject("ADODB.Recordset")
    html = " "

    If UCase(table) = UCase("workingmonth") Then
        sql = " Select DISTINCT " & itmID & " FROM " & table & " Order by " & itmID & " Desc"
    Else
        sql = " Select DISTINCT " & itmID & " FROM " & table
    End If

    html = html & " <div class='period-selector df f1 plr25 mlr25'> "
    html = html & "     <span class='plr10 t2 mw-115'>" & itmName & "</span>"
    html = html & "     <input class='select-value1' type='text' id='filterInput" & ID & "' placeholder='" & itmName & "' />"
    html = html & "     <select id='" & ID & "' class='select-value1'>"
    If UCase(table) = UCase("ItemStore") Or UCase(table) = UCase("WorkingMonth") Then
        html = html & "       <option value='All'>All " & itmName & "s</option>"
    Else
        html = html & "       <option value=''>All " & itmName & "s</option>"
    End If

    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Do Until rst.EOF
            val = rst.fields("" & itmID & "")
            valName = GetComboName("" & table & "", rst.fields("" & itmID & ""))

            If UCase(CStr(val)) = UCase(sup) Or UCase(CStr(val)) = UCase(mth) Then
                html = html & "         <option class='opt-sty' value='" & CStr(val) & "' id='" & CStr(val) & "' selected >" & valName & "</option>"
            Else
                html = html & "         <option class='opt-sty' value='" & CStr(val) & "' id='" & CStr(val) & "'>" & valName & "</option>"
            End If

            rst.MoveNext
            html = html & response.flush
        Loop
    End If

    html = html & "     </select>"

    html = html & " <script>" & vbCrLf
    html = html & "   var filterInput" & ID & " = document.getElementById('filterInput" & ID & "');"
    html = html & "   var filterDropdown" & ID & " = document.getElementById('" & ID & "');"
    html = html & "   filterInput" & ID & ".addEventListener('input', function () {"
    html = html & "     var filterValue = filterInput" & ID & ".value.toLowerCase();"
    html = html & "     for (var i = 0; i < filterDropdown" & ID & ".options.length; i++) {"
    html = html & "       var optionText = filterDropdown" & ID & ".options[i].text.toLowerCase();"
    html = html & "       var optionValue = filterDropdown" & ID & ".options[i].value.toLowerCase();"
    html = html & "       if ("
    html = html & "         optionText.includes(filterValue) ||"
    html = html & "         optionValue.includes(filterValue)"
    html = html & "       ) {"
    html = html & "         filterDropdown" & ID & ".selectedIndex = i;"
    html = html & "         break;"
    html = html & "       }"
    html = html & "     }"
    html = html & "   });"
    html = html & "</script>" & vbCrLf
    html = html & " </div>"

    rst.Close
    Set rst = Nothing
    DetailSelectorFilter = html
End Function

Sub TableEX()

    response.write "        <div class='p15 df dnp pt0'>"
    response.write "          <button id='exportCSV' class='btsty bs'>"
    response.write "            <span>Export To CSV &nbsp;</span>"
    response.write "            <i class='fa-solid fa-file-csv'></i>"
    response.write "          </button>"
    response.write "          <button id='exportExcel' class='btsty bs'>"
    response.write "            <span>Export To Excel &nbsp;</span>"
    response.write "            <i class='fa-solid fa-file-excel'></i>"
    response.write "          </button>"
    response.write "          <button id='printButton' class='btsty bs'>"
    response.write "            <span>Print Table &nbsp;</span>"
    response.write "            <i class='fa-solid fa-print'></i>"
    response.write "          </button>"
    response.write "        <div class='df f1 plr25 mlr25'>"
    response.write "          <div class='t2 plr10'>Search Table</div>"
    response.write "          <input class='select-value1' type='text' id='searchInput' placeholder='search..'/>"
    response.write "        </div>"
    sTb2 = "ItemPurOrder"
    If HasAccessRight(uName, "frm" & sTb2, "New") Then
        response.write "        <div class='df f1 plr25 mlr25'> " & Addlnk("Add New Purchase Order", "wpgItemPurOrder.asp?PageMode=AddNew", "bp") & "</div>"
    End If
    response.write "        </div>"

End Sub

Sub StylesAdded()

    response.write "     <style>"
    response.write "       :root {"
    response.write "         --bc: white;"
    response.write "         --bp: #0994de;"
    response.write "         --br: #ffd54a;"
    response.write "         --brhv: #ffc400;"
    response.write "         --bphv: #087ab8;"
    response.write "         --bss: #05ab5d;"
    response.write "         --bsshv: #198754;"
    response.write "         --bwrn: #ff3f52;"
    response.write "         --bwrnhv: #dc3545;"
    response.write "         --bs: #939393;"
    response.write "         --bshv: #6c757d;"
    response.write "         --bsty-primary: #005f92;"
    response.write "         --bsty-alert: #dc3545;"
    response.write "         --bsty-alert-hover: #ff6776;"
    response.write "         --bsty-primary-hover: #0994de;"
    response.write "         --bsty-success: #198754;"
    response.write "         --bsty-success-hover: #05ab5d;"
    response.write "         --bsty-sec: #ffc400;"
    response.write "         --bsty-sec-hover: #ffd54a;"
    response.write "       }"
    response.write "       .holder {"
    response.write "         display: flex;"
    response.write "       }"
    response.write "       .btsty {"
    response.write "         width: fit-content;"
    response.write "         padding: 4px;"
    response.write "         margin: 0px 0px 0px 5px;"
    response.write "         border-radius: 4px;"
    response.write "         cursor: pointer;"
    response.write "         text-align: center;"
    response.write "         white-space: nowrap;"
    response.write "         border: none;"
    response.write "         font-size: 13.33px;"
    response.write "         text-decoration: none;"
    response.write "       }"
    response.write "       .pm {"
    response.write "         color: var(--bsty-primary);"
    response.write "       }"
    response.write "       .pm:hover {"
    response.write "         text-decoration: underline;"
    response.write "         color: var(--bsty-primary-hover);"
    response.write "       }"
    response.write "       .rd {"
    response.write "         color: var(--bsty-sec);"
    response.write "       }"
    response.write "       .rd:hover {"
    response.write "         text-decoration: underline;"
    response.write "         color: var(--bsty-sec-hover);"
    response.write "       }"
    response.write "       .bnm:hover {"
    response.write "         text-decoration: underline;"
    response.write "       }"
    response.write "       .wrn {"
    response.write "         color: var(--bsty-alert);"
    response.write "       }"
    response.write "       .wrn:hover {"
    response.write "         color: var(--bsty-alert-hover);"
    response.write "         text-decoration: underline;"
    response.write "       }"
    response.write "       .scc {"
    response.write "         color: var(--bsty-success);"
    response.write "       }"
    response.write "       .scc:hover {"
    response.write "         color: var(--bsty-success-hover);"
    response.write "         text-decoration: underline;"
    response.write "       }"
    response.write "       .bp {"
    response.write "         background-color: var(--bp);"
    response.write "         color: var(--bc);"
    response.write "         margin: 0px 5px 0px 5px;"
    response.write "         padding: 4px 12px;"
    response.write "       }"
    response.write "       .br {"
    response.write "         background-color: var(--br);"
    response.write "         color: var(--bc);"
    response.write "         margin: 0px 5px 0px 5px;"
    response.write "         padding: 4px 12px;"
    response.write "       }"
    response.write "       .obp {"
    response.write "         box-shadow: 0px 0px 1px 1px var(--bp);"
    response.write "         color: var(--bphv);"
    response.write "         margin: 0px 5px 0px 5px;"
    response.write "         padding: 4px 12px;"
    response.write "       }"
    response.write "       .obp:hover {"
    response.write "         background-color: var(--bp);"
    response.write "         color: var(--bc);"
    response.write "       }"
    response.write "       .bp:hover {"
    response.write "         background-color: var(--bphv);"
    response.write "       }"
    response.write "       .bss {"
    response.write "         background-color: var(--bss);"
    response.write "         color: var(--bc);"
    response.write "         margin: 0px 5px 0px 5px;"
    response.write "         padding: 4px 12px;"
    response.write "       }"
    response.write "       .obss {"
    response.write "         box-shadow: 0px 0px 1px 1px var(--bss);"
    response.write "         color: var(--bsshv);"
    response.write "         margin: 0px 5px 0px 5px;"
    response.write "         padding: 4px 12px;"
    response.write "       }"
    response.write "       .obss:hover {"
    response.write "         background-color: var(--bss);"
    response.write "         color: var(--bc);"
    response.write "       }"
    response.write "       .bss:hover {"
    response.write "         background-color: var(--bsshv);"
    response.write "       }"
    response.write "       .br:hover {"
    response.write "         background-color: var(--brhv);"
    response.write "       }"
    response.write "       .bwrn {"
    response.write "         background-color: var(--bwrn);"
    response.write "         color: var(--bc);"
    response.write "         margin: 0px 5px 0px 5px;"
    response.write "         padding: 4px 12px;"
    response.write "       }"
    response.write "       .obwrn {"
    response.write "         box-shadow: 0px 0px 1px 1px var(--bwrn);"
    response.write "         color: var(--bwrnhv);"
    response.write "         margin: 0px 5px 0px 5px;"
    response.write "         padding: 4px 12px;"
    response.write "       }"
    response.write "       .obwrn:hover {"
    response.write "         background-color: var(--bwrn);"
    response.write "         color: var(--bc);"
    response.write "       }"
    response.write "       .bwrn:hover {"
    response.write "         background-color: var(--bwrnhv);"
    response.write "       }"
    response.write "       .bs {"
    response.write "         background-color: var(--bs);"
    response.write "         color: var(--bc);"
    response.write "         margin: 0px 5px 0px 5px;"
    response.write "         padding: 4px 12px;"
    response.write "       }"
    response.write "       .obs {"
    response.write "         box-shadow: 0px 0px 1px 1px var(--bs);"
    response.write "         color: var(--bshv);"
    response.write "         margin: 0px 5px 0px 5px;"
    response.write "         padding: 4px 12px;"
    response.write "       }"
    response.write "       .obs:hover {"
    response.write "         background-color: var(--bshv);"
    response.write "         color: var(--bc);"
    response.write "       }"
    response.write "       .bs:hover {"
    response.write "         background-color: var(--bshv);"
    response.write "       }"
    response.write "       .bold {"
    response.write "         font-weight: 600;"
    response.write "       }"
    response.write "       .aln {"
    response.write "         padding: 0px;"
    response.write "         margin: 0px;"
    response.write "       }"
    response.write "       .f-end {"
    response.write "         justify-content: flex-end;"
    response.write "       }"
    response.write "       .alf-end {"
    response.write "         align-items: flex-end;"
    response.write "       }"
    response.write "       .table-holder::-webkit-scrollbar {"
    response.write "         width: 5px;"
    response.write "         height: 5px;"
    response.write "       }"
    response.write "       .table-holder::-webkit-scrollbar-track {"
    response.write "         background: #efefef;"
    response.write "       }"
    response.write "       .table-holder::-webkit-scrollbar-thumb {"
    response.write "         background: #b4b4b4;"
    response.write "       }"
    response.write "       .table-holder::-webkit-scrollbar-thumb:hover {"
    response.write "         background: #888;"
    response.write "       }"
    response.write "       .table-holder::-webkit-scrollbar-button:hover {"
    response.write "         background: #888;"
    response.write "       }"
    response.write "       .tbl {"
    response.write "         width: 100%;"
    response.write "         border-collapse: collapse;"
    response.write "       }"
    response.write "       .tbl1 {"
    response.write "         width: 100%;"
    response.write "         border-collapse: collapse;"
    response.write "       }"
    response.write "       .fh {"
    response.write "         padding: 7px 16px;"
    response.write "         font-size: 12px;"
    response.write "         position: sticky;"
    response.write "         top: 0;"
    response.write "         z-index: 9;"
    response.write "         background-color: #f9fafb;"
    response.write "       }"
    response.write "       .fh-s {"
    response.write "         padding: 7px 16px;"
    response.write "         font-size: 12px;"
    response.write "         position: sticky;"
    response.write "         top: 0;"
    response.write "         z-index: 9;"
    response.write "       }"
    response.write "       .tbl-b tr:first-child td {"
    response.write "         border-top: 1px solid #ddd;"
    response.write "       }"
    ' response.write "       .tbl-Sub tr td:first-child {"
    ' response.write "         border-left: 1px solid #ddd;"
    ' response.write "       }"
    response.write "       .tbl-Sub tr td {"
    response.write "         padding: 5px !important;"
    response.write "         font-size: 13px !important;"
    response.write "         background: none;"
    response.write "       }"
    ' response.write "       .tbl-Sub tr td:last-child {"
    ' response.write "         border-right: 1px solid #ddd;"
    ' response.write "       }"
    response.write "       .tbl-Sub tr th {"
    response.write "         font-size: 12px;"
    response.write "         padding: 5px 10px;"
    response.write "       }"
    response.write "       .tbl-b tr:last-child td {"
    response.write "         border-bottom: none;"
    response.write "       }"
    response.write "       .tbl-b tr:last-child td:first-child {"
    response.write "         border-bottom-left-radius: 8px;"
    response.write "       }"
    response.write "       .tbl-b tr:last-child td:last-child {"
    response.write "         border-bottom-right-radius: 8px;"
    response.write "       }"
    response.write "       .tbl-h tr:first-child th:first-child {"
    response.write "         border-top-left-radius: 8px;"
    response.write "       }"
    response.write "       .tbl-h tr:first-child th:last-child {"
    response.write "         border-top-right-radius: 8px;"
    response.write "       }" 'tbl-Sub
    response.write "       .tbl-b tr td {"
    response.write "         padding: 10px 8px;"
    response.write "         border-bottom: 1px solid #ddd;"
    response.write "         background-color: white;"
    response.write "         font-size: 14px;"
    response.write "         vertical-align: top;"
    response.write "       }"
    response.write "       .tbl1 tbody tr td {"
    response.write "         padding: 0px;"
    response.write "         border: none !important;"
    response.write "         background-color: white;"
    response.write "         font-size: 14px;"
    response.write "       }"
    response.write "       .tbl-main {"
    response.write "         border: 1px solid #ddd;"
    response.write "         border-radius: 8px;"
    response.write "       }"
    response.write "       .main-content {"
    response.write "         height: calc(100% - 80px);"
    response.write "         overflow: auto;"
    response.write "         padding-bottom: 20px;"
    response.write "         background-color: #f9fafb;"
    response.write "       }"
    response.write "       .Select-name {"
    response.write "         font-size: 24px;"
    response.write "         font-weight: 600;"
    response.write "       }"
    response.write "       .datastat {"
    response.write "         font-size: 12px;"
    response.write "         width: fit-content;"
    response.write "         padding: 3px 8px;"
    response.write "         border-radius: 8px;"
    response.write "         white-space: nowrap;"
    response.write "       }"
    response.write "     </style>"
    response.write "     <style>"
    response.write "       * {"
    response.write "         margin: 0;"
    response.write "         padding: 0;"
    response.write "         font-family: 'Segoe UI';"
    response.write "       }"
    response.write "       body {"
    response.write "         background-color: #f9fafb;"
    response.write "       }"
    response.write "       .df {"
    response.write "         display: flex;"
    response.write "         align-items: center;"
    response.write "       }"
    response.write "       .rlt {"
    response.write "         position: relative;"
    response.write "       }"
    response.write "       .rt0 {"
    response.write "         right: 0;"
    response.write "       }"
    response.write "       .abs {"
    response.write "         position: absolute;"
    response.write "       }"
    response.write "       .clsbtt {"
    response.write "         top: -28px;"
    response.write "         right: 0;"
    response.write "       }"
    response.write "       .fs {"
    response.write "         align-items: flex-start;"
    response.write "       }"
    response.write "       .fc {"
    response.write "         flex-direction: column;"
    response.write "       }"
    response.write "       .sb {"
    response.write "         justify-content: space-between;"
    response.write "       }"
    response.write "       .jc {"
    response.write "         justify-content: center;"
    response.write "       }"
    response.write "       .j-sb {"
    response.write "         justify-content: space-between;"
    response.write "       }"
    response.write "       .f1 {"
    response.write "         flex: 1;"
    response.write "       }"
    response.write "       .h100 {"
    response.write "         height: 100%;"
    response.write "       }"
    response.write "       .cp {"
    response.write "         cursor: pointer;"
    response.write "       }"
    response.write "       .p10 {"
    response.write "         padding: 10px;"
    response.write "       }"
    response.write "       .plr5 {"
    response.write "         padding-inline: 5px;"
    response.write "       }"
    response.write "       .plr10 {"
    response.write "         padding-inline: 10px;"
    response.write "       }"
    response.write "       .plr15 {"
    response.write "         padding-inline: 15px;"
    response.write "       }"
    response.write "       .plr25 {"
    response.write "         padding-inline: 25px;"
    response.write "       }"
    response.write "       .p15 {"
    response.write "         padding: 15px;"
    response.write "       }"
    response.write "       .p25 {"
    response.write "         padding: 25px;"
    response.write "       }"
    response.write "       .pb10 {"
    response.write "         padding-bottom: 10px;"
    response.write "       }"
    response.write "       .pb15 {"
    response.write "         padding-bottom: 15px;"
    response.write "       }"
    response.write "       .pb25 {"
    response.write "         padding-bottom: 25px;"
    response.write "       }"
    response.write "       .pt10 {"
    response.write "         padding-top: 10px;"
    response.write "       }"
    response.write "       .pt15 {"
    response.write "         padding-top: 15px;"
    response.write "       }"
    response.write "       .pt25 {"
    response.write "         padding-top: 25px;"
    response.write "       }"
    response.write "       .m10 {"
    response.write "         margin: 10px;"
    response.write "       }"
    response.write "       .mlr10 {"
    response.write "         margin-inline: 10px;"
    response.write "       }"
    response.write "       .mlr15 {"
    response.write "         margin-inline: 15px;"
    response.write "       }"
    response.write "       .mlr25 {"
    response.write "         margin-inline: 25px;"
    response.write "       }"
    response.write "       .m15 {"
    response.write "         margin: 15px;"
    response.write "       }"
    response.write "       .m25 {"
    response.write "         margin: 25px;"
    response.write "       }"
    response.write "       .w100 {"
    response.write "         width: 100%;"
    response.write "       }"
    response.write "       .t1 {"
    response.write "         font-size: 18px;"
    response.write "         font-weight: 600;"
    response.write "       }"
    response.write "       .st1 {"
    response.write "         font-size: 16px;"
    response.write "       }"
    response.write "       .sblg {"
    response.write "         width: 185px;"
    response.write "       }"
    response.write "       .st1s {"
    response.write "         font-size: 15px;"
    response.write "         color: #a1a0a0;"
    response.write "       }"
    response.write "       .t2 {"
    response.write "         font-size: 16px;"
    response.write "         font-weight: 500;"
    response.write "       }"
    response.write "       .t3 {"
    response.write "         font-size: 16px;"
    response.write "         font-weight: 400;"
    response.write "       }"
    response.write "       .st2 {"
    response.write "         font-size: 14px;"
    response.write "       }"
    response.write "       .st3 {"
    response.write "         font-size: 14px;"
    response.write "         font-weight: 300;"
    response.write "       }"
    response.write "       .st2s {"
    response.write "         font-size: 13px;"
    response.write "         color: #a1a0a0;"
    response.write "       }"
    response.write "       .st3s {"
    response.write "         font-size: 13px;"
    response.write "         color: #a1a0a0;"
    response.write "         font-weight: 300;"
    response.write "       }"
    response.write "       .aprv {"
    response.write "         background-color: #def7ec;"
    response.write "         font-weight: 600;"
    response.write "         color: #046c4e;"
    response.write "       }"
    response.write "       .dnd {"
    response.write "         background-color: #fde8e8;"
    response.write "         font-weight: 600;"
    response.write "         color: #c01e1e;"
    response.write "       }"
    response.write "       .pnd {"
    response.write "         background-color: #feecdc;"
    response.write "         font-weight: 600;"
    response.write "         color: #b43403;"
    response.write "       }"
    response.write "       .expd {"
    response.write "         background-color: #f4f5f7;"
    response.write "         font-weight: 600;"
    response.write "         color: #24262d;"
    response.write "       }"
    response.write "       .c-aprv {"
    response.write "         color: #046c4e;"
    response.write "       }"
    response.write "       .c-dnd {"
    response.write "         color: #c01e1e;"
    response.write "       }"
    response.write "       .c-pnd {"
    response.write "         color: #dd761b;"
    response.write "       }"
    response.write "       .c-expd {"
    response.write "         color: #24262d;"
    response.write "       }"
    response.write "       .icn {"
    response.write "         width: 45px;"
    response.write "         height: 45px;"
    response.write "         border-radius: 50%;"
    response.write "         font-size: 20px;"
    response.write "       }"
    response.write "       .top-dt {"
    response.write "         border: 1px solid #ddd;"
    response.write "         border-radius: 8px;"
    response.write "         background-color: white;"
    response.write "       }"
    response.write "       .tbl-b1 tr:first-child td {"
    response.write "         border-top: 1px solid #ddd;"
    response.write "       }"
    response.write "       .tbl-b1 tr:last-child td {"
    response.write "         border-bottom: none;"
    response.write "       }"
    response.write "       .tbl-b1 tr:last-child td:first-child {"
    response.write "         border-bottom-left-radius: 8px;"
    response.write "       }"
    response.write "       .tbl-b1 tr:last-child td:last-child {"
    response.write "         border-bottom-right-radius: 8px;"
    response.write "       }"
    response.write "       .tbl-b1 tr td {"
    response.write "         padding: 12px 16px;"
    response.write "         border-bottom: 1px solid #ddd;"
    response.write "         background-color: white;"
    response.write "         font-size: 14px;"
    response.write "         vertical-align: top;"
    response.write "       }"
    response.write "       .p0 {"
    response.write "         padding: 0px !important;"
    response.write "       }"
    response.write "       .pt0 {"
    response.write "         padding-top: 0px !important;"
    response.write "       }"
    response.write "       .pr0 {"
    response.write "         padding-right: 0px !important;"
    response.write "       }"
    response.write "       .pr5 {"
    response.write "         padding-right: 5px !important;"
    response.write "       }"
    response.write "       .pl0 {"
    response.write "         padding-left: 0px !important;"
    response.write "       }"
    response.write "       .pb0 {"
    response.write "         padding-bottom: 0px !important;"
    response.write "       }"
    response.write "       td .p5 {"
    response.write "         padding: 5px !important;"
    response.write "       }"
    response.write "       .fh1 {"
    response.write "         padding: 5px 10px;"
    response.write "         font-size: 12px;"
    response.write "         background-color: #f9fafb;"
    response.write "       }"
    response.write "       .meal-tbl tr td {"
    response.write "         padding: 2px;"
    response.write "       }"
    response.write "       .meal-hld {"
    response.write "         flex-grow: 1;"
    response.write "       }"
    response.write "       .meal-hld:first-child {"
    response.write "         background-color: #fff6b689;"
    response.write "       }"
    response.write "       .meal-hld:nth-child(2) {"
    response.write "         background-color: #ffdcb686;"
    response.write "       }"
    response.write "       .meal-hld:last-child {"
    response.write "         background-color: #ffc6b689;"
    response.write "       }"
    response.write "       .meal-hld:first-child {"
    response.write "         border-bottom-left-radius: 8px;"
    response.write "         border-top-left-radius: 8px;"
    response.write "       }"
    response.write "       .meal-hld:last-child {"
    response.write "         border-bottom-right-radius: 8px;"
    response.write "         border-top-right-radius: 8px;"
    response.write "       }"
    response.write "       .menu {"
    response.write "         background-color: white;"
    response.write "         align-items: stretch;"
    response.write "         width: fit-content;"
    response.write "         border-radius: 8px;"
    response.write "         box-shadow: 0px 0px 3px 1px #ddd;"
    response.write "       }"
    response.write "       .opt-sty {"
    response.write "         font-size: 12px;"
    response.write "         background: #eeeeee;"
    response.write "       }"
    response.write " .select-value1 {"
    response.write "   padding-inline: 8px;"
    response.write "   padding-top: 4px;"
    response.write "   padding-bottom: 4px;"
    response.write "   font-size: 12px !important;"
    ' response.Write "   display: block;"
    response.write "   box-sizing: border-box;"
    response.write "   border-radius: 4px;"
    response.write "   border: 1px solid #dddddd;"
    response.write "   outline: none;"
    response.write "   background: #f8f8f8;"
    response.write "   transition: background 0.2s, border-color 0.2s;"
    ' response.Write "   max-width: 16%;"
    response.write "   min-width: 16%;"
    response.write "   text-align: left;"
    response.write " }"
    response.write " .select-value1:focus {"
    response.write "   border-color: #a8d9e4;"
    response.write "   background: #ffffff;"
    response.write " }"
    response.write "  .pt0 {"
    response.write "    padding-top: 0px !important;"
    response.write "  }"
    response.write "  .pr0 {"
    response.write "    padding-right: 0px !important;"
    response.write "  }"
    response.write "  .pl0 {"
    response.write "    padding-left: 0px !important;"
    response.write "  }"
    response.write "  .pb0 {"
    response.write "    padding-bottom: 0px !important;"
    response.write "  }"
    response.write "  td .p5 {"
    response.write "    padding: 5px !important;"
    response.write "  }"
    response.write "  .p5 {"
    response.write "    padding: 5px !important;"
    response.write "  }"
    response.write "  .fh1 {"
    response.write "    padding: 5px 10px;"
    response.write "    font-size: 12px;"
    response.write "    background-color: #f9fafb;"
    response.write "  }"
    response.write "  .dn {"
    response.write "    display: none !important;"
    response.write "    border: none !important;"
    response.write "  }"
    response.write " .custom-tooltip {"
    response.write "   position: relative;"
    response.write "   display: inline-block;"
    response.write "   cursor: pointer;"
    response.write " }"
    response.write " .tooltip-text {"
    response.write "   display: none;"
    response.write "   position: absolute;"
    response.write "   left: 10px;"
    response.write "   background-color: #afafaf;"
    response.write "   color: #fff;"
    response.write "   padding: 5px;"
    response.write "   border-radius: 6px;"
    response.write " }"
    response.write " .p24 {"
    response.write "   padding: 2px 4px;"
    response.write " }"
    response.write "       .tbl-b tr:hover td {"
    response.write "         background-color: #fafafa;"
    response.write "       }"
    ' response.Write "        th[data-sort='asc']::after {"
    ' response.Write "           content: ' ?';"
    ' response.Write "           font-size: 12px;"
    ' response.Write "           line-height: 1;"
    ' response.Write "         }"
    ' response.Write "         th[data-sort='desc']::after {"
    ' response.Write "           content: ' ?';"
    ' response.Write "           font-size: 12px;"
    ' response.Write "           line-height: 1;"
    ' response.Write "         }"
    response.write " .hld-abs {"
    response.write "    bottom: 20px;"
    response.write "    z-index: 999;"
    response.write "    display: none;"
    response.write "  }"
    response.write "  .prc-hld {"
    response.write "    background-color: #f9fafb;"
    response.write "    border-radius: 12px;"
    response.write "    box-shadow: 1px 1px 15px 0px #dfdfdf;"
    response.write "  }"
    response.write "  .tbls tr td {"
    response.write "    border: none !important;"
    response.write "    background: none !important;"
    response.write "  }"
    response.write "  .rlt-act:hover .hld-abs {"
    response.write "    display: block;"
    response.write "  }"
    response.write "         @media print {"
    response.write "           .dnp {"
    response.write "             display: none !important;"
    response.write "           }"
    response.write "         }"
    response.write "    .min-top {"
    response.write "      right: 20px;"
    response.write "      top: 0;"
    response.write "      visibility: hidden;"
    response.write "      opacity: 0;"
    response.write "      width: 30px;"
    response.write "      height: 30px;"
    response.write "      background: #ededed;"
    response.write "      border-radius: 50%;"
    response.write "      color: #959494;"
    response.write "      transition: all 0.7s ease;"
    response.write "    }"
    response.write "    .t-case:hover .min-top,"
    response.write "    .min-top:hover .t-case-i {"
    response.write "      visibility: visible;"
    response.write "      opacity: 1;"
    response.write "    }"
    response.write "    .icnn {"
    response.write "      transition: all 0.5s ease;"
    response.write "    }"
    response.write "    .t-case-i {"
    response.write "      top: -11px;"
    response.write "      right: 13px;"
    response.write "      background: #efefef;"
    response.write "      padding: 2px 4px;"
    response.write "      border-radius: 5px;"
    response.write "      color: #595959;"
    response.write "      font-size: 13px;"
    response.write "      visibility: hidden;"
    response.write "      opacity: 0;"
    response.write "      transition: all 1s ease;"
    response.write "    }"
    response.write "    .min-top:hover {"
    response.write "      background: #d1d1d1;"
    response.write "      color: #474747;"
    response.write "    }"
    response.write "    .act-t {"
    response.write "      visibility: visible;"
    response.write "      opacity: 1;"
    response.write "    }"
    response.write "    .act-t > .fa-angle-up {"
    response.write "      rotate: 180deg;"
    response.write "    }"
    response.write "    .t-case-m {"
    response.write "      padding-top: 0;"
    response.write "      padding-bottom: 2px;"
    response.write "    }"
    response.write "    .dn {"
    response.write "      max-width: 0;"
    response.write "      overflow: hidden;"
    response.write "      display: none;"
    response.write "    }"
    response.write "    .mw-115 {"
    response.write "      min-width: 115px;"
    response.write "    }"
    response.write "    .mw-300 {"
    response.write "      min-width: 300px;"
    response.write "    }"
    response.write "       .t-right {"
    response.write "         text-align: right;"
    response.write "         padding-right: 8px;"
    response.write "       }"
    response.write "       .t-left {"
    response.write "         text-align: left;"
    response.write "       }"
    response.write "       .t-center {"
    response.write "         text-align: center;"
    response.write "       }"
    response.write "       .fh-s {"
    response.write "         padding: 7px 16px;"
    response.write "         font-size: 12px;"
    response.write "         font-weight: 600;"
    response.write "       }"
    response.write "   .licnt {"
    response.write "     border-radius: 8px;"
    response.write "     border: 1px solid #ddd;"
    response.write "     background: whitesmoke;"
    response.write "     top: 0px;"
    response.write "     z-index: 99;"
    response.write "   }"
    response.write "   .noti{"
    response.write "     border-radius: 10px;"
    response.write "     width: fit-content;"
    response.write "   }"
    response.write "   .revl{"
    response.write "     max-width: 0px;"
    response.write "     padding: 0px;"
    response.write "     transition: all 0.8s ease;"
    response.write "     overflow: hidden;"
    response.write "     white-space: nowrap;"
    response.write "   }"
    response.write "   .noti:hover .revl{"
    response.write "     max-width: 160px;"
    response.write "     padding: 3px 8px;"
    response.write "   }"
    response.write "     </style>"

End Sub


Sub TableEXScp()

    response.write " <script>"
    response.write "   const prevButton = document.getElementById('prevPage');"
    response.write "   const nextButton = document.getElementById('nextPage');"
    response.write "   const currentPageDisplay = document.getElementById('currentPage');"
    response.write "   const table = document.querySelector('.tbl');"
    response.write "   const rows = Array.from(table.querySelectorAll('tbody tr'));"
    response.write "   const searchInput = document.getElementById('searchInput');"
    response.write "   const printButton = document.getElementById('printButton');"
    response.write "   const rowsPerPage = 10;"
    response.write "   const tbody = table.querySelector('tbody');"
    response.write "   const ths = Array.from(table.querySelectorAll('th'));"
    response.write "   let currentPage = 1;"
    ' response.write "   function displayRows() {"
    ' response.write "     const startIndex = (currentPage - 1) * rowsPerPage;"
    ' response.write "     const endIndex = startIndex + rowsPerPage;"
    ' response.write "     rows.forEach((row, index) => {"
    ' response.write "       if (index >= startIndex && index < endIndex) {"
    ' response.write "         row.style.display = 'table-row';"
    ' response.write "       } else {"
    ' response.write "         row.style.display = 'none';"
    ' response.write "       }"
    ' response.write "     });"
    ' response.write "   }"
    ' response.write "   displayRows();"
    ' response.write "   prevButton.addEventListener('click', () => {"
    ' response.write "     if (currentPage > 1) {"
    ' response.write "       currentPage--;"
    ' response.write "       currentPageDisplay.textContent = currentPage;"
    ' response.write "       displayRows();"
    ' response.write "     }"
    ' response.write "   });"
    ' response.write "   nextButton.addEventListener('click', () => {"
    ' response.write "     if (currentPage < Math.ceil(rows.length / rowsPerPage)) {"
    ' response.write "       currentPage++;"
    ' response.write "       currentPageDisplay.textContent = currentPage;"
    ' response.write "       displayRows();"
    ' response.write "     }"
    ' response.write "   });"
    response.write "   function filterRows() {"
    response.write "     const query = searchInput.value.toLowerCase();"
    response.write "     rows.forEach((row) => {"
    response.write "       const text = row.innerText.toLowerCase();"
    response.write "       if (text.includes(query)) {"
    response.write "         row.style.display = 'table-row';"
    response.write "       } else {"
    response.write "         row.style.display = 'none';"
    response.write "       }"
    response.write "     });"
    response.write "   }"
    response.write "   searchInput.addEventListener('input', filterRows);"
    response.write "   function getDataType(columnIndex) {"
    response.write "     const rows = Array.from(tbody.querySelectorAll('tr'));"
    response.write "     for (const row of rows) {"
    response.write "       const cellValue = row.querySelector("
    response.write "         `td:nth-child(${columnIndex + 1})`"
    response.write "       ).textContent;"
    response.write "       if (!isNaN(parseFloat(cellValue))) {"
    response.write "         return 'number';"
    response.write "       }"
    response.write "     }"
    response.write "     return 'string';"
    response.write "   }"
    response.write "   function sortTable(columnIndex) {"
    response.write "     if (ths[columnIndex].classList.contains('nosort')) {"
    response.write "       return;"
    response.write "     }"
    response.write "     const dataType = getDataType(columnIndex);"
    response.write "     const isAscending ="
    response.write "       ths[columnIndex].getAttribute('data-sort') === 'asc';"
    response.write "     const rows = Array.from(tbody.querySelectorAll('tr'));"
    response.write "     rows.sort((a, b) => {"
    response.write "       const cellA = a.querySelector("
    response.write "         'td:nth-child(' + (columnIndex + 1) + ')'"
    response.write "       ).textContent;"
    response.write "       const cellB = b.querySelector("
    response.write "         'td:nth-child(' + (columnIndex + 1) + ')'"
    response.write "       ).textContent;"
    response.write "       if (dataType === 'number') {"
    response.write "         return isAscending"
    response.write "           ? parseFloat(cellA) - parseFloat(cellB)"
    response.write "           : parseFloat(cellB) - parseFloat(cellA);"
    response.write "       } else {"
    response.write "         return isAscending"
    response.write "           ? cellA.localeCompare(cellB)"
    response.write "           : cellB.localeCompare(cellA);"
    response.write "       }"
    response.write "     });"
    response.write "     rows.forEach((row) => tbody.appendChild(row));"
    response.write "     ths.forEach((th) => th.setAttribute('data-sort', 'none'));"
    response.write "     ths[columnIndex].setAttribute("
    response.write "       'data-sort',"
    response.write "       isAscending ? 'desc' : 'asc'"
    response.write "     );"
    response.write "   }"
    response.write "   ths.forEach((th, index) => {"
    response.write "     th.addEventListener('click', () => {"
    response.write "       sortTable(index);"
    response.write "     });"
    response.write "   });"
    response.write "   function exportTableToExcel() {"
    response.write "     const table = document.querySelector('table');"
    response.write "     const rows = Array.from(table.querySelectorAll('tr'));"
    response.write "     let excelContent = '<table>';"
    response.write "     rows.forEach((row) => {"
    response.write "       const cells = Array.from(row.querySelectorAll('td, th')); "
    response.write "       excelContent += '<tr>';"
    response.write "       cells.forEach((cell, index) => {"
    response.write "         if ("
    response.write "           !table"
    response.write "             .querySelector(`th:nth-child(${index + 1}`)"
    response.write "             .classList.contains('ignore')"
    response.write "         ) {"
    response.write "           excelContent += `<td>${cell.textContent}</td>`;"
    response.write "         }"
    response.write "       });"
    response.write "       excelContent += '</tr>';"
    response.write "     });"
    response.write "     excelContent += '</table>';"
    response.write "     const excelBlob = new Blob([excelContent], {"
    response.write "       type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',"
    response.write "     });"
    response.write "     const link = document.createElement('a');"
    response.write "     link.href = URL.createObjectURL(excelBlob);"
    response.write "     link.setAttribute('download', 'table_data.xls');"
    response.write "     link.click();"
    response.write "   }"
    response.write "   function exportTableToCSV() {"
    response.write "     const table = document.querySelector('table');"
    response.write "     const rows = Array.from(table.querySelectorAll('tr'));"
    response.write "     let csvContent = '';"
    response.write "     rows.forEach((row) => {"
    response.write "       const cells = Array.from(row.querySelectorAll('td, th')); "
    response.write "       csvContent += cells.map((cell) => `'${cell.textContent}'`).join(',');"
    response.write "       csvContent += '\n';"
    response.write "     });"
    response.write "     const csvBlob = new Blob([csvContent], {"
    response.write "       type: 'text/csv;charset=utf-8;',"
    response.write "     });"
    response.write "     const link = document.createElement('a');"
    response.write "     link.href = URL.createObjectURL(csvBlob);"
    response.write "     link.setAttribute('download', 'table_data.csv');"
    response.write "     link.click();"
    response.write "   }"
    response.write "   const exportExcelButton = document.getElementById('exportExcel');"
    response.write "   exportExcelButton.addEventListener('click', exportTableToExcel);"
    response.write "   const exportCSVButton = document.getElementById('exportCSV');"
    response.write "   exportCSVButton.addEventListener('click', exportTableToCSV);"
    response.write "   function printTable() {"
    response.write "     window.print();"
    response.write "   }"
    response.write "   printButton.addEventListener('click', printTable);"
    response.write " </script>"

End Sub

Sub ScpAdded()
    response.write "<script>"
    response.write " const cntTp = document.querySelector('.cnt');"
    response.write " const cntTptg = document.querySelector('.min-top');"
    response.write " const cntTc = document.querySelector('.t-case');"
    response.write " const cntTci = document.querySelector('.t-case-i');"
    response.write " cntTptg.addEventListener('click', () => {"
    response.write "   cntTptg.classList.toggle('act-t');"
    response.write "   if (cntTptg.classList.contains('act-t')) {"
    response.write "     cntTp.classList.add('dn');"
    response.write "     cntTc.classList.add('t-case-m');"
    response.write "     cntTci.textContent = 'Maximize';"
    response.write "   } else {"
    response.write "     cntTp.classList.remove('dn');"
    response.write "     cntTc.classList.remove('t-case-m');"
    response.write "     cntTci.textContent = 'Minimize';"
    response.write "   }"
    response.write " });"
    response.write "</script>"

    response.write " <script>"
    response.write " function openLinkInPopup(url) {"
    response.write "     window.open(url, '_self');"
    response.write "   }"

    response.write "  const mainElements = document.querySelectorAll('[class^=""btsty bs showsrv""]');"
    response.write "  const mainElements1 = document.querySelectorAll('[class^=""clsbtt clsbtn""]');"
    response.write "  const subElements = document.querySelectorAll('[class^=""df p5 abs licnt dn srvlist""]');"
    response.write "  mainElements.forEach((mainElement, index) => {"
    response.write "    mainElement.addEventListener('click', () => {"
    response.write "      for (let i = 0; i < subElements.length; i++) {"
    response.write "        if (i !== index) {subElements[i].classList.add('dn');}"
    response.write "      }"
    response.write "      subElements[index].classList.remove('dn');"
    response.write "    });"
    response.write "  });"
    response.write "  mainElements1.forEach((mainElement1, index) => {"
    response.write "    mainElement1.addEventListener('click', () => {"
    response.write "      subElements[index].classList.add('dn');"
    response.write "    });"
    response.write "  });"

    response.write "  const bmainElements = document.querySelectorAll('[class^=""btsty bs bdshow""]');"
    response.write "  const bmainElements1 = document.querySelectorAll('[class^=""clsbtt bdbtn""]');"
    response.write "  const bsubElements = document.querySelectorAll('[class^=""df p5 abs licnt rt0 dn bdlist""]');"
    response.write "  bmainElements.forEach((bmainElement, index) => {"
    response.write "    bmainElement.addEventListener('click', () => {"
    response.write "      for (let i = 0; i < bsubElements.length; i++) {"
    response.write "        if (i !== index) {bsubElements[i].classList.add('dn');}"
    response.write "      }"
    response.write "      bsubElements[index].classList.remove('dn');"
    response.write "    });"
    response.write "  });"
    response.write "  bmainElements1.forEach((bmainElement1, index) => {"
    response.write "    bmainElement1.addEventListener('click', () => {"
    response.write "      bsubElements[index].classList.add('dn');"
    response.write "    });"
    response.write "  });"

    response.write " function openPopup(linkUrl, width, height) {"
    response.write " var left = (window.innerWidth - width) / 2;"
    response.write " var top = (window.innerHeight - height) / 2;"
    response.write " window.open(linkUrl, 'PopupWindow', 'width=' + width + ', height=' + height + ', left=' + left + ', top=' + top);"
    response.write " }"

    response.write " </script>"

End Sub

'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>
