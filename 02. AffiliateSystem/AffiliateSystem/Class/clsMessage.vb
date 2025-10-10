Imports DevExpress.Web

Public Class clsMessage

    Public Enum MsgType
        InformationMessage = 0
        ErrorMessage = 1
        ErrorMessageFromVS = 2
    End Enum

    Public Sub DisplayMessage(ByRef pLabel As ASPxEditors.ASPxLabel, ByVal pMsgID As String, ByVal pMsgType As MsgType, Optional ByVal pMessage As String = "")
        Dim ls_Message As String = ""

        Select Case pMsgType
            Case 0 '# INFORMATION MESSAGE
                pLabel.ForeColor = Drawing.Color.Blue

                Select Case pMsgID
                    Case "1001"
                        ls_Message = "Data Saved Successfully!"
                    Case "1002"
                        ls_Message = "Data Updated Successfully!"
                    Case "1003"
                        ls_Message = "Data Deleted Successfully!"
                    Case "1004"
                        ls_Message = "Data Copied successfully!"
                    Case "1005"
                        ls_Message = "Data Uploaded successfully!"
                    Case "1006"
                        ls_Message = "Data Approved successfully!"
                    Case "1007"
                        ls_Message = "Affiliate Approved successfully!"
                    Case "1008"
                        ls_Message = "Affiliate Final Approval successfully!"
                    Case "1009"
                        ls_Message = "Data UnApprove Successfully!"
                    Case "1010"
                        ls_Message = "Send Good Receiving to Supplier Successfully!"
                    Case "1011"
                        ls_Message = "Affiliate UnApprove successfully!"
                    Case "1012"
                        ls_Message = "Send Kanban to Supplier Successfully!"
                    Case "2001"
                        ls_Message = "No data that you want to search!"
                    Case "2002"
                        ls_Message = "Period does not match!"
                    Case "2003"
                        ls_Message = "Cost Center is not valid!"
                    Case "2004"
                        ls_Message = "Account Code is not valid!"

                    Case "3001"
                        ls_Message = "Process completed!"
                    Case "3002"
                        ls_Message = "Allocation process successful!"
                    Case "3003"
                        ls_Message = "No data to allocate!"

                End Select

            Case 1 '# WARNING/ERROR MESSAGE
                pLabel.ForeColor = Drawing.Color.Red

                Select Case pMsgID
                    Case "5001"
                        ls_Message = "Please Input Firm Quantity!"
                    Case "5002"
                        ls_Message = "Please Input Forecast Month 1!"
                    Case "5003"
                        ls_Message = "Please Input Forecast Month 2!"
                    Case "5004"
                        ls_Message = "Please Input Forecast Month 3!"
                    Case "5005"
                        ls_Message = "Firm Qty must be same or multiple of the MOQ!"
                    Case "5006"
                        ls_Message = "E.T.A Qty must be same with Firm Qty!"
                    Case "5007"
                        ls_Message = "Daily Qty must be same or multiple of the Qty/Box!"
                    Case "5008"
                        ls_Message = "Forecast Month 1 must be same or multiple of the MOQ!"
                    Case "5009"
                        ls_Message = "Forecast Month 2 be same or multiple of the MOQ!"
                    Case "5010"
                        ls_Message = "Forecast Month 3 be same or multiple of the MOQ!"
                    Case "5011"
                        ls_Message = "Not found price!"
                    Case "5012"
                        ls_Message = "This PO No. Alraedy exist!"
                    Case "5013"
                        ls_Message = "This PO Revision No. Alraedy exist!"
                    Case "5014"
                        ls_Message = "This PO No. Alraedy create PO Revision!"
                    Case "5015"
                        ls_Message = "PO Qty Revision can't bigger than Original PO Qty!"
                    Case "5016"
                        ls_Message = "Please choose the file!"
                    Case "6001"
                        ls_Message = "Invalid User Name or Password!"
                    Case "6010"
                        ls_Message = "Please select the data first!"
                    Case "6011"
                        ls_Message = "No data edited!"
                    Case "6012"
                        ls_Message = "Please input %% first!"
                    Case "6014"
                        ls_Message = "Can't process. Please press Check button first!"
                    Case "6015"
                        ls_Message = "Can't process. Because still error found!"
                    Case "6016"
                        ls_Message = "Can't insert the same Item Code with Parent "
                    Case "6017"
                        ls_Message = "No data edited!"
                    Case "6018"
                        ls_Message = "Data Already Exist!"
                    Case "6019"
                        ls_Message = "Can't delete this Record, Record used by another Application "
                    Case "6020"
                        ls_Message = "Location is not valid!"
                    Case "6021"
                        ls_Message = "No data to process!"
                    Case "6022"
                        ls_Message = "Price must be LOWER than 999.99 !"
                    Case "6023"
                        ls_Message = "Amount must be LOWER than 9,999,999,999,999,999.99 %%!"
                    Case "6024"
                        ls_Message = "Employment Group is not valid!"
                    Case "6025"
                        ls_Message = "Currency is not valid!"
                    Case "6026"
                        ls_Message = "%% is invalid numeric value!"
                    Case "6027"
                        ls_Message = "Can't Update/Delete. Data already Approved !"
                    Case "6028"
                        ls_Message = "Can't process. Please Input Exchange Rate !"
                    Case "6029"
                        ls_Message = "Please Select Supplier !"
                    Case "6030"
                        ls_Message = "Period must be same with kanban date !"
                    Case "6031"
                        ls_Message = "Kanban Qty must be lower than Remaining PO Qty ! "

                    Case "6101"
                        ls_Message = "Start Date Depreciation can't less than Period !"
                    Case "6102"
                        ls_Message = "Over Budget in this P/R. Please Input Remarks Purpose !"
                    Case "6103"
                        ls_Message = "Over Budget in this P/R. Are you sure want to continue save the data? "


                    Case "7001"
                        ls_Message = "Value can't greater than %% !"
                    Case "7002"
                        ls_Message = "Current Password is not Correct!"
                    Case "7003"
                        ls_Message = "Please choose Cost Center Code First!"
                    Case "7004"
                        ls_Message = "Please input User ID First!"
                    Case "7005"
                        ls_Message = "Please input Full Name First!"
                    Case "7006"
                        ls_Message = "Please input Password First!"
                    Case "7007"
                        ls_Message = "Please input Confirmation Password First!"
                    Case "7008"
                        ls_Message = "New Password and Confirm Password doesn't match, please try again!"
                    Case "7009"
                        ls_Message = "Data Already Final Approval"
                    Case "7010"
                        ls_Message = "Please input Request Qty first!"
                    Case "7011"
                        ls_Message = "Please Approve Data First !"
                    Case "7012"
                        ls_Message = "Please Select Data First !"

                    Case "8001"
                        ls_Message = "Please input PONo. you want to export."
                End Select

            Case Else '# ERROR MESSAGE FROM SYSTEM
                pLabel.ForeColor = Drawing.Color.Red
                ls_Message = pMessage

        End Select

        pLabel.Text = "[" & pMsgID & "] " & ls_Message
    End Sub

End Class
