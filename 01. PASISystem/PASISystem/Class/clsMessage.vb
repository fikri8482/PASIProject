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
                        ls_Message = "Data Copied successfully!"
                    Case "1008"
                        ls_Message = "Send to Supplier successfully!"
                    Case "1009"
                        ls_Message = "Approve Data successfully!"
                    Case "1010"
                        ls_Message = "Send Good Receiving to Supplier Successfully!"
                    Case "1011"
                        ls_Message = "Send Tally Successfully !"
                    Case "1012"
                        ls_Message = "Send Shipping Successfully !"
                    Case "1013"
                        ls_Message = "Send Invoice Successfully !"
                    Case "1014"
                        ls_Message = "PO Split successfully!"
                    Case "1015"
                        ls_Message = "PO Cancel successfully!"
                    Case "1016"
                        ls_Message = "Recovery Data successfully!"
                    Case "2001"
                        ls_Message = "No data that you want to search!"
                    Case "2007"
                        ls_Message = "No data that you want to download!"
                    Case "2002"
                        ls_Message = "Period does not match!"
                    Case "2003"
                        ls_Message = "Cost Center is not valid!"
                    Case "2004"
                        ls_Message = "Account Code is not valid!"
                    Case "2005"
                        ls_Message = "Send E.D.I Data Successfully !"
                    Case "2006"
                        ls_Message = "Send Shipping Instruction Data Successfully !"
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
                        ls_Message = "Can't process. Affiliate ID already used in other screen"
                    Case "5002"
                        ls_Message = "Can't process. Supplier ID already used in other screen"
                    Case "5003"
                        ls_Message = "Can't process. Part No. already used in other screen"
                    Case "5004"
                        ls_Message = "Can't process. This Data already used in other screen"
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
                        ls_Message = "Cannot Update/Delete, Data already Approved !"
                    Case "6028"
                        ls_Message = "Can't process. Please Input Exchange Rate !"

                    Case "6029"
                        ls_Message = "Qty/Box must be multiply of the Carton Qty and Delivery Qty"

                    Case "6030"
                        ls_Message = "Cut Of Date must be same with Period!"

                    Case "6040"
                        ls_Message = "Can't select different Affiliate!"

                    Case "6041"
                        ls_Message = "Can't select different Forwarder!"

                    Case "6101"
                        ls_Message = "Start Date Depreciation can't less than Period !"
                    Case "6102"
                        ls_Message = "Over Budget in this P/R. Please Input Remarks Purpose !"
                    Case "6103"
                        ls_Message = "Over Budget in this P/R. Are you sure want to continue save the data? "
                    Case "6104"
                        ls_Message = "Box No. not Register on system !"
                    Case "6105"
                        ls_Message = "Consignee Code Already Exist!"

                    Case "6201"
                        ls_Message = "Price with this Incoterm not found, please check price master again!"

                    Case "7001"
                        ls_Message = "Value can't greater than %% !"
                    Case "7002"
                        ls_Message = "Current Password is not Correct!"
                    Case "7003"
                        ls_Message = "Please choose Affiliate ID First!"
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
                        ls_Message = "Page not found!"
                    Case "7010"
                        ls_Message = "Please input Request Qty first!"
                    Case "7011"
                        ls_Message = "Please Approve Data First !"
                    Case "7012"
                        ls_Message = "Please Select Data First !"
                    Case "7013"
                        ls_Message = "Can't Process different Affiliate !"
                    Case "7014"
                        ls_Message = "Qty Can't bigger than Receiving Qty !"
                    Case "7015"
                        ls_Message = "Please choose Good or Defect Qty !"
                    Case "7016"
                        ls_Message = "Invalid Box No. !"
                    Case "7009"
                        ls_Message = "Budget Year To not more than Budget Year From"

                    Case "8001"
                        ls_Message = "Budget period already closed, cannot be updated!"
                    Case "8002"
                        ls_Message = "Qty Box DOPASI and PLPASI Don't Match !"
                    Case "8003"
                        ls_Message = "Send E.D.I Failed : " & pMessage
                End Select

            Case Else '# ERROR MESSAGE FROM SYSTEM
                pLabel.ForeColor = Drawing.Color.Red
                ls_Message = pMessage

        End Select

        pLabel.Text = "[" & pMsgID & "] " & ls_Message
    End Sub

End Class
