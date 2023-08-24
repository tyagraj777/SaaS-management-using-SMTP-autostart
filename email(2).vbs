Dim sch, cdoConfig, cdoMessage
sch = "http://schemas.microsoft.com/cdo/configuration/"
Set cdoConfig = CreateObject("CDO.Configuration")
With cdoConfig.Fields
.Item(sch & "sendusing") = 1 ' cdoSendUsingPort
.Item(sch & "smtpserver") = "127.0.0.1"
'    .Item(sch & "smtpserverport") = 25
.update
End With
Set cdoMessage = CreateObject("CDO.Message")
With cdoMessage
Set .Configuration = cdoConfig
.From = "admin@qispqchatupmgr1.na.qualcomm.com"
.To = "qis.hyd.ce.qchat@qualcomm.com"
.Subject = "Alert@QISUpload_Master_Shutdown_Initiated..!!!"
.TextBody = "Host-admin@qispqchatupmgr1.na.qualcomm.com IP-10.46.95.168"
'.AddAttachment "c:\images\myimage.jpg"
.Send
End With
Set cdoMessage = Nothing
Set cdoConfig = Nothing
'MsgBox "Email Sent"