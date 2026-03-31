
'<<--BEGIN_CODE_SEGMENT_PRINTHEADER-->>
Dim patientID
patientID = Trim(Request.queryString("patientID"))

response.write "<!DOCTYPE html>"
response.write "<html lang=""en"">"

response.write "<head> "
'response.write "  <title>Vehicle Management</title>"
' response.write "<link href=""https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600&display=swap"" rel=""stylesheet"">"
' response.write "<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css"" integrity=""sha384-rbsA2VBKQhggwzxH7pPCaAqO46MgnOM80zW1RWuH61DGLwZJEdK2Kadq2F9CUG65"" crossorigin=""anonymous"">"
' response.write "<link href=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.css"" rel=""stylesheet"">"
' response.write "<link href=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.css"" rel=""stylesheet"">"
  response.write Glob_GetBootstrap5()
  response.write Glob_GetIconFontAwesome()
 response.write "</head>"


'response.write Glob_GetBootstrap5()
'response.write Glob_GetIconFontAwesome()

SetPageVariable "AutoHidePrintControl", "Yes"
ShowIncomingInvoices


'If UCase(jSchd) = UCase("SystemAdmin") Or UCase(jSchd) = UCase(uName) Or UCase(GetComboNameFld("SystemUser", uName, "StaffID")) = UCase("STF001") Then
SetPageAlerts
'End If
AddAutoRefreshJs

Sub ShowIncomingInvoices()
    Dim sql, rst, html, flagDetail, fLen, href, blink, foceGray, dt, style

    AddPageJS

    sql = "select top 150 Patient.PatientName, PatientFlag2.PatientFlag2Name, PatientFlag2.PatientFlag2ID, Gender.GenderName, PatientFlag2.FlagDetail2, PatientFlag2.FlagDetail1 "
    sql = sql & " , PatientFlag2.FlagInfo1,PatientFlag2.FlagDetail4, PatientFlag2.FlagInfo3, PatientFlag2.SystemUserID, PatientFlag2.EntryDate, JobSchedule.JobScheduleName, PatientFlag2.PatientFlagModeID"
    sql = sql & " , PatientFlag2.PatientID, Receipt.ReceiptID, PatientFlag2.PatientFlagStatusID, PatientFlag2.FlagValue1, PatientFlag2.FirstDayID "
    sql = sql & " from PatientFlag2 "
    sql = sql & " left join Receipt on Receipt.ReceiptInfo1=PatientFlag2.PatientFlag2ID"
    sql = sql & " left join Patient on Patient.PatientID=PatientFlag2.PatientID"
    sql = sql & " left join Gender on Gender.GenderID=Patient.GenderID"
    sql = sql & " left join JobSchedule on JobSchedule.JobScheduleID=PatientFlag2.JobScheduleID "
    ' sql = sql & " where PatientFlag2.PatientFlagModeID='P003' and PatientFlag2.PatientFlagTypeID='P003' "
    sql = sql & " where PatientFlag2.PatientFlagModeID IN ('P003','P004') and PatientFlag2.PatientFlagTypeID='P003' " ''Invoice, Kitchen
    If Len(patientID) > 0 Then
        sql = sql & " and PatientFlag2.PatientID='" & patientID & "' "
    End If
    ' sql = sql & " and PatientFlag2.BranchID='" & brnch & "' "
    ' sql = sql & " order by PatientFlag2.EntryDate desc;" ''@bless - 6 May 2024 //show unpaid invoice first
    ' sql = sql & " order by Receipt.ReceiptInfo1, PatientFlag2.EntryDate desc;"
    sql = sql & " order by PatientFlag2.FirstDayID desc, Receipt.ReceiptInfo1, PatientFlag2.EntryDate desc;"

    Set rst = CreateObject("ADODB.RecordSet")

    html = html & "<style>"
    html = html & "     .invoice-listing{width:98vw;margin-top:20px;} "
    html = html & "     .force-gray, .force-gray *{color:gray!important;} "
    html = html & "     .invoice-listing th{background-image:linear-gradient(to bottom, #e3ffff, #e3ffff, #cae4dd, #e3ffff, #e3ffff);color:#1e5b5b;height:40px;top:0;position:sticky;} "
    html = html & "     .invoice-listing, .invoice-listing>tbody>tr>td, .invoice-listing>thead>tr>th{ "
    html = html & "         border:1px solid silver;border-collapse:collapse;padding:5px 10px;text-transform:uppercase;font-size:12px;"
    html = html & "     }"
    html = html & "     .invoice-listing .blink{ "
    html = html & "         animation: blink-animation 2s steps(50, start) infinite alternate 2s;-webkit-animation: blink-animation 2s steps(50, start) infinite alternate 2s;"
    html = html & "     }"
    html = html & "     @keyframes blink-animation {"
    html = html & "       to {"
    html = html & "         visibility: hidden;"
    html = html & "       }"
    html = html & "     }"
    html = html & "     @-webkit-keyframes blink-animation {"
    html = html & "       to {"
    html = html & "         visibility: hidden;"
    html = html & "       }"
    html = html & "     }"
    html = html & "     option{padding:3px;}"
    html = html & "</style>"
    AddPageCSS2
    AddPageJS2

    response.write vbCrLf & "    <div class=""ii"">"
    response.write vbCrLf & "        <span style=""color: black;"">Search:</span>"
    response.write vbCrLf & "        <div class=""search-b2"">"
    response.write vbCrLf & "            <div class=""search-box"">"
    response.write vbCrLf & "                <input type=""text"" id=""inpPatientID"" name=""inpPatientID"" value=""" & patientID & """>"
    response.write vbCrLf & "                <a href=""javascript:window.open('wpgFindLargeTableKey.asp?SelectInfo=inpPatientID||Patient||PatientID&SrcTable=','_blank','height=300px,width=500px')"">Find</a>"
    response.write vbCrLf & "            </div>"
    response.write vbCrLf & "            <button onclick=""searchParam()"">Process</button>"
    response.write vbCrLf & "        </div>"
    response.write vbCrLf & "    </div>"




    html = html & "<table id='vehicleTable' class='invoice-listing'>"
    html = html & " <thead>"
    html = html & "     <tr>"
    html = html & "         <th>No.</th>"
    html = html & "         <th>Invoice<br/>No.</th>"
    html = html & "         <th>Patient ID</th>"
    html = html & "         <th>Patient Name</th>"
    html = html & "         <th>Gender</th>"
    html = html & "         <th>Details</th>"
    html = html & "         <th>Invoice Amount</th>"
    html = html & "         <th>Department</th>"
    html = html & "         <th>By</th>"
    html = html & "         <th>Date</th>"
    html = html & "         <th>Control</th>"
    html = html & "     </tr>"
    html = html & " </thead>"

    fLen = 80
    dt = Now()
    rst.open qryPro.FltQry(sql), conn, 3, 4
    If rst.RecordCount > 0 Then
        rst.movefirst
        Do While Not rst.EOF
            blink = ""
            foceGray = ""
            style = ""
            pat = rst.fields("PatientID")

            flagDetail = Replace(rst.fields("FlagDetail2"), vbCrLf, "<br/>")
            If Len(flagDetail) > fLen Then
                flagDetail = Left(flagDetail, fLen) & " ..."
            End If

            If (DateDiff("n", rst.fields("EntryDate"), dt) < 100) Then
                'recents
                If (IsNull(rst.fields("ReceiptID")) Or (rst.fields("ReceiptID") = "")) And (rst.fields("PatientFlagStatusID") = "P001") Then
                    'new/active
                    blink = "blink"
                    'style = "style='color:#2196F3;font-weight:bold;'"
                    style = "style='color:#a94442;font-weight:bold;'"
                ElseIf Len(rst.fields("ReceiptID")) > 0 Then
                    'foceGray = "force-gray"
                    'style = "style='color:#3c763d;'"
                End If
            ElseIf rst.fields("PatientFlagStatusID") <> "P001" Then
                'inactive
                foceGray = "force-gray"
            End If

            html = html & "     <tr class='" & foceGray & " " & blink & "' " & style & ">"
            html = html & "         <td>" & rst.AbsolutePosition & "</td>"
            html = html & "         <td>" & rst.fields("PatientFlag2ID") & "</td>"
            If UCase(pat) = UCase("P1") Then
                flgMd = rst.fields("PatientFlagModeID")
                ' html = html & "         <td>" & rst.fields("PatientName") & "</td>"
                If UCase(flgMd) = UCase("P004") Then
                    html = html & "         <td>" & rst.fields("FlagDetail1") & " [" & GetComboName("PatientFlagMode", flgMd) & "]</td>"
                Else
                    html = html & "         <td>" & rst.fields("PatientName") & "</td>"
                End If
                html = html & "         <td>" & rst.fields("PatientFlag2Name") & "</td>"
                html = html & "         <td>" & "N/A" & "</td>"
            Else
                html = html & "         <td>" & rst.fields("PatientID") & "</td>"
                html = html & "         <td>" & rst.fields("PatientName") & "</td>"
                html = html & "         <td>" & rst.fields("GenderName") & "</td>"
            End If
            html = html & "         <td>" & flagDetail & "</td>"
            html = html & "         <td><b>" & FormatNumber(rst.fields("FlagValue1"), 2) & "</b></td>"
            html = html & "         <td>" & rst.fields("JobScheduleName") & "</td>"
            html = html & "         <td>" & GetComboName("Staff", GetComboNameFld("SystemUser", rst.fields("SystemUserID"), "StaffID")) & "</td>"
            html = html & "         <td style='text-transform:none;'>" & GetHowLong(rst.fields("EntryDate")) & "</td>"

            If Len(rst.fields("ReceiptID")) > 0 Then
                ' href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=ReceiptSlip2&PositionForTableName=Receipt&ReceiptID=" & rst.fields("ReceiptID")
                href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PaymentReceipt&PositionForTableName=Receipt&ReceiptID=" & rst.fields("ReceiptID")
                html = html & " <td>" & GetLink(href, "Print Receipt", "#3c763d") & "</td>"
            ElseIf rst.fields("PatientFlagStatusID") = "P001" Then
                href = "wpgReceipt.asp?PageMode=AddNew&PullUpData=PatientID||" & rst.fields("PatientID")
                href = href & "&InvoiceNo=" & rst.fields("PatientFlag2ID")
                html = html & " <td>" & GetLink(href, "Issue Receipt", "")
                href = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=InstallmentAgree&PositionForTableName=WorkingDay&WorkingDayID=&invoiceno=" & rst.fields("PatientFlag2ID")
                If UCase(rst.fields("FlagInfo3")) = "CONSULTREVIEW" And UCase(rst.fields("FlagDetail4")) <> "INSTALLMENT-PAYMENT" Then
                    html = html & "<br><br><span >" & GetLink(href, "Create Installment", "forestgreen")
                End If
                html = html & "</td>"
            Else
                html = html & " <td></td>"
            End If

            html = html & "     </tr>"
            rst.movenext
        Loop
        rst.Close
        html = html & " </tbody>"
    End If

    Set rst = Nothing
    html = html & "</table>"
    response.write html
End Sub

Function GetHowLong(dt)
    Dim tmpDt, ot, dDiff, hrs
    If Not IsDate(dt) Then
        GetHowLong = dt
    Else
        tmpDt = Now()
        dDiff = DateDiff("n", dt, tmpDt)

        If dDiff >= 0 Then
            If dDiff < 2 Then
                ot = "just now"
            ElseIf dDiff < 60 Then
                ot = CInt(dDiff) & " mins ago"
            Else
                hrs = CInt(dDiff / 60)
                If hrs < 24 Then
                    ot = hrs & " hours ago"
                ElseIf hrs >= 24 And hrs < 48 Then
                    ot = "yesterday"
                Else
                    ot = CInt(hrs / 24) & " days ago"
                End If
            End If
        Else
            ot = FormatDate(dt)
        End If
        GetHowLong = ot
    End If
End Function

Sub AddAutoRefreshJs()
    Dim html
    html = "<script>"
    html = html & " function autoRefresh(){ "
    html = html & "     if(new Date().getTime() - last_ac_time >= 20000){"
    html = html & "         window.location.reload(); "
    html = html & "     }"
    html = html & "     else{"
    html = html & "         setTimeout(autoRefresh, 20000);"
    html = html & "     }"
    html = html & " }"
    html = html & " var last_ac_time = new Date().getTime();"
    'html = html & " document.addEventListener('mousemove', function(){last_ac_time = new Date().getTime()});"
    html = html & " document.addEventListener('click', function(){last_ac_time = new Date().getTime()});"
    html = html & " setTimeout(autoRefresh, 20000);"
    html = html & " </script>"
    response.write html
End Sub

Sub AddPageJS()
    Dim html
    html = "<script>"
    html = html & " function  processDateChange(select){ "
    html = html & "     let url_search = new URL(window.location.href);"
    html = html & "     url_search.searchParams.set('AppointDayID', select.options[select.selectedIndex].value);"
    html = html & "     window.location.href = url_search;"
    html = html & " }"
    html = html & " function  processSpecialistChange(select){ "
    html = html & "     let url_search = new URL(window.location.href);"
    html = html & "     url_search.searchParams.set('SpecialistID', select.options[select.selectedIndex].value);"
    html = html & "     window.location.href = url_search;"
    html = html & " }"
    html = html & " function openPopup(anc){"
    html = html & "     let win=window.open(anc.dataset.href, '_blank', 'resizeable=yes,scrollbars=yes,width=820,height=560,status=yes');"
    html = html & "     "
    html = html & "     let intvl = setInterval(function(){"
    html = html & "         if(win.closed !== false){"
    html = html & "             clearInterval(intvl);"
    html = html & "             window.location.reload();"
    html = html & "          }"
    html = html & "     }, 200);"
    html = html & "}"
    html = html & "</script>"
    response.write html
End Sub

Function GetLink(href, linkText, lnkColor)
    Dim html, defColor
    defColor = IIF(Trim(lnkColor) = "", "#2196F3", lnkColor)

    html = "<div style='display:inline-block;color:" & defColor & ";font-weight:bold;text-transform:none;'>"
    html = html & "<span data-href=""" & href & """ style='/*text-decoration:underline;*/cursor:pointer;' "
    html = html & "onclick=""openPopup(this)"""
    html = html & ">" & linkText & "</span>"
    html = html & "</div>"
    GetLink = html
End Function
Function IIF(expression, trueVal, falseVal)
    If expression = True Then
        IIF = trueVal
    Else
        IIF = falseVal
    End If
End Function


'Sub SetPageAlerts()
'    response.write Glob_GetBootstrapToastAlertHeader("")
'    Set rst = CreateObject("ADODB.Recordset")
'    dtNow = Now()
'
'    'Discharged patient ready for printing bills alert
'    minsAgo = 60 * 6 ''6 hours
'    vDt1 = FormatDateDetail(DateAdd("n", (-1 * minsAgo), dtNow))
'    vDt2 = FormatDateDetail(dtNow)
'
'    sql = "select Count(a.visitationID) as Cnt, a.TransProcessValID "
'    sql = sql & " from Admission as a, Visitation as v "
'    sql = sql & " where v.visitationID=a.visitationID and v.BranchID='" & brnch & "'  "
'    sql = sql & " and a.AdmissionStatusID IN ('A008') And v.InitialVisitationID<>'SUB'  "
'    ' sql = sql & " and (a.MainValue1 > 0) and a.DischargeDate between '" & vDt1 & "' and '" & vDt2 & "'  "
'    sql = sql & " And a.DischargeDate between '" & vDt1 & "' and '" & vDt2 & "'  "
'    sql = sql & " group by a.TransProcessValID "
'    ' Response.write sql
'
'
'    With rst
'        rst.open qryPro.FltQry(sql), conn, 3, 4
'        If rst.RecordCount > 0 Then
'            rst.MoveFirst
'            ' response.write Glob_GetBootstrapToastAlertHeader("")
'            Set tOption = server.CreateObject("Scripting.Dictionary")
'            cnt = rst.fields("Cnt")
'            alertText = cnt & " patient"
'            If IsNumeric(cnt) And cnt > 1 Then alertText = alertText & "s"
'            alertText = alertText & " non-OPD clients ready to pay bills for " & GetComboName("Branch", brnch)
'
'            tOption.Add "close", True
'            tOption.Add "icon", True
'            tOption.Add "delay", 60
'
'            tOption.Add "title", "Discharged Patients"
'            tOption.Add "subtitle", "Ready to Pay Bills"
'            tOption.Add "button1", "See Details"
'            url = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=DischargeForBillingList&PositionForTableName=WorkingDay"
'            url = url & "&DisplayType=RecentDischarges&RecentDischargeDate=" & vDt1
'            tOption.Add "button1Url", url
'            lnkCnt = lnkCnt + 1
'            response.write Glob_GetBootstrapToastAlert("Success", alertText, tOption, lnkCnt)
'            response.flush
'            Set tOption = Nothing
'        End If
'        rst.Close
'    End With
'    Set rst = Nothing
'
'    response.write Glob_GetBootstrapToastAlertFooter()
'End Sub

Sub SetPageAlerts()
    response.write Glob_GetBootstrapToastAlertHeader("")
    Set rst = CreateObject("ADODB.Recordset")
    dtNow = Now()

    'Discharged patient ready for printing bills alert
    minsAgo = 60 * 6 ''6 hours
    vDt1 = FormatDateDetail(DateAdd("n", (-1 * minsAgo), dtNow))
    vDt2 = FormatDateDetail(dtNow)

    sql = "select Count(a.visitationID) as Cnt, a.TransProcessValID "
    sql = sql & " from Admission as a, Visitation as v "
    sql = sql & " where v.visitationID=a.visitationID and v.BranchID='" & brnch & "'  "
    sql = sql & " and a.AdmissionStatusID IN ('A008') And v.InitialVisitationID<>'SUB'  "
    ' sql = sql & " and (a.MainValue1 > 0) and a.DischargeDate between '" & vDt1 & "' and '" & vDt2 & "'  "
    sql = sql & " And a.DischargeDate between '" & vDt1 & "' and '" & vDt2 & "'  "
    sql = sql & " group by a.TransProcessValID "



    With rst
        rst.open qryPro.FltQry(sql), conn, 3, 4
        If rst.RecordCount > 0 Then
            rst.movefirst
            ' response.write Glob_GetBootstrapToastAlertHeader("")
            Set tOption = server.CreateObject("Scripting.Dictionary")
            cnt = rst.fields("Cnt")
            alertText = cnt & " patient"
            If IsNumeric(cnt) And cnt > 1 Then alertText = alertText & "s"
            alertText = alertText & " non-OPD clients ready to pay bills for " & GetComboName("Branch", brnch)

            tOption.Add "close", True
            tOption.Add "icon", True
            tOption.Add "delay", 60

            tOption.Add "title", "Discharged Patients"
            tOption.Add "subtitle", "Ready to Pay Bills"
            tOption.Add "button1", "See Details"
            url = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=DischargeForBillingList&PositionForTableName=WorkingDay"
            url = url & "&DisplayType=RecentDischarges&RecentDischargeDate=" & vDt1
            tOption.Add "button1Url", url
            lnkCnt = lnkCnt + 1
            'response.write Glob_GetBootstrapToastAlert("Success", alertText, tOption, lnkCnt)
            response.flush
            Set tOption = Nothing
        End If
        rst.Close
    End With
    Set rst = Nothing



       Dim rstR, sqlR, ot
    ot = False
    Set rstR = CreateObject("ADODB.Recordset")




    dtNow = Now()

    'Unvalidated receipt  alert
    minsAgo = 60 * 6 ''6 hours
    vDt1 = FormatDateDetail(DateAdd("n", (-1 * minsAgo), dtNow))
    vDt2 = FormatDateDetail(dtNow)

'   sqlR = "SELECT * FROM Receipt WHERE SystemUserID='" & usr & "' AND CustomerTypeID='C109' AND ReceiptAmount2 <= 0.1 "
    sqlR = "SELECT COUNT(DISTINCT r.receiptid) as cnt FROM Receipt r "
    sqlR = sqlR & " Join Admission ad ON r.patientid = ad.patientid"
    sqlR = sqlR & " WHERE r.CustomerTypeID='C109' AND r.ReceiptAmount2 <= 0.1  "
    sqlR = sqlR & " AND r.WorkingDayID >= 'DAY20231201' "
    sqlR = sqlR & " And ad.Workingdayid > 'DAY20231201' AND ad.AdmissionstatusID= 'A008'"


    'response.write sqlR
    ' Response.write sql


    With rstR
        rstR.open qryPro.FltQry(sqlR), conn, 3, 4
        If rstR.RecordCount > 0 Then
            rstR.movefirst
            ' response.write Glob_GetBootstrapToastAlertHeader("")
            Set tOption = server.CreateObject("Scripting.Dictionary")
            cnt = rstR.fields("cnt")
            alertText = cnt & " Inpatient"
            If IsNumeric(cnt) And cnt > 1 Then alertText = alertText & "s"
            alertText = alertText & " reciept(s) is pending validation for " & GetComboName("Branch", brnch)

            tOption.Add "close", True
            tOption.Add "icon", True
            tOption.Add "delay", 60

            tOption.Add "title", "Receipt Validation"
            tOption.Add "subtitle", "Pending Validation"
            tOption.Add "button1", "See Details"
            url = "wpgPrtPrintLayoutAll.asp?PrintLayoutName=PendingInpatientRecieptValidation&PositionForTableName=WorkingDay"

            tOption.Add "button1Url", url
            lnkCnt = lnkCnt + 1
            response.write Glob_GetBootstrapToastAlert("danger", alertText, tOption, lnkCnt)
            response.flush
            Set tOption = Nothing
        End If
        rst.Close
    End With
    Set rst = Nothing


    response.write Glob_GetBootstrapToastAlertFooter()
End Sub

Sub AddPageCSS2()
    response.write vbCrLf & "    <style>"
    response.write vbCrLf & "        .ii{"
    response.write vbCrLf & "            display: flex;"
    response.write vbCrLf & "            justify-content: end;"
    response.write vbCrLf & "            align-items: center;"
    response.write vbCrLf & "            padding: 0 20px;"
    response.write vbCrLf & "            gap: 10px;"
    response.write vbCrLf & "            margin:5px 0;"
    response.write vbCrLf & "            font-size:12px;"
    response.write vbCrLf & "        }"
    response.write vbCrLf & "        .search-b2{"
    response.write vbCrLf & "            display: flex;"
    response.write vbCrLf & "            box-sizing: border-box;"
    response.write vbCrLf & "            font-family: inherit;"
    response.write vbCrLf & "        }"
    response.write vbCrLf & "        .search-box{"
    response.write vbCrLf & "            border: 1px solid grey;"
    response.write vbCrLf & "            padding: 4px 5px;"
    response.write vbCrLf & "        }"
    response.write vbCrLf & "        .search-box a{"
    response.write vbCrLf & "            text-decoration: none;"
    response.write vbCrLf & "            color: white;"
    response.write vbCrLf & "            background-color: #0078D4;"
    response.write vbCrLf & "            padding: 2px 10px;"
    response.write vbCrLf & "            border-radius: 4px;"
    response.write vbCrLf & "        }"
    response.write vbCrLf & "        .search-box input{"
    response.write vbCrLf & "            border: none;"
    response.write vbCrLf & "            outline: none;"
    response.write vbCrLf & "            font-family: inherit;"
    response.write vbCrLf & "        }"
    response.write vbCrLf & "        .search-b2 button{"
    response.write vbCrLf & "            background-color: forestgreen;"
    response.write vbCrLf & "            color: white;"
    response.write vbCrLf & "            border: none;"
    response.write vbCrLf & "            font-family: inherit;"
    response.write vbCrLf & "        }"
    response.write vbCrLf & "    </style>"
End Sub

Sub AddPageJS2()
    response.write "<script>"
    response.write "function searchParam(){"
    response.write "const value = document.getElementById('inpPatientID').value;"
    response.write " const url = new URL(window.location.href);"
    response.write " url.searchParams.set('patientID', value);"
    response.write " window.location.href = url;"
    response.write "}"
    response.write "</script>"
End Sub

'response.write "<script src=""https://code.jquery.com/jquery-3.6.0.min.js""></script>"
'response.write "<script src=""https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/js/bootstrap.bundle.min.js"" integrity=""sha384-kenU1KFdBIe4zVF0s0G1M5b4hcpxyD9F7jL+jjXkk+Q2h455rYXK/7HAuoJl+0I4"" crossorigin=""anonymous""></script>"
'response.write "<script src=""https://cdn.datatables.net/v/dt/dt-1.13.6/datatables.min.js""></script>"
'response.write "<script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js""></script>"
'response.write "<script src=""https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js""></script>"
'response.write "<script src=""https://cdn.datatables.net/v/bs5/jszip-3.10.1/dt-1.13.6/b-2.4.1/b-html5-2.4.1/b-print-2.4.1/datatables.min.js""></script>"

'response.write "<script>"
'response.write "        var table = $('#vehicleTable').DataTable({"
'response.write "                       paging: false, "
'response.write "                       pageLength: 200, "
'response.write "                       lengthChange: false,"
'response.write "                       lengthMenu: [100, 300, 500, 700],"
'response.write "        });"
'response.write ""
'response.write "        table.buttons().container().appendTo('#vehicleTable_wrapper .col-md-6:eq(0)');"
'
'response.write "</script>"
'<<--END_CODE_SEGMENT_PRINTHEADER-->>
'>
'>
'>
'>
'>
'<<--BEGIN_CODE_SEGMENT_PRINTFOOTER-->>

'<<--END_CODE_SEGMENT_PRINTFOOTER-->>