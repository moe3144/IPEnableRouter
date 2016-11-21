Imports Microsoft.Win32
Imports System.ServiceProcess
Imports System.Management

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ''' Enable IP Routing in Registry by updateing the proper key.
        Dim key As RegistryKey = My.Computer.Registry.LocalMachine
        Dim subkey As RegistryKey
        subkey = key.OpenSubKey("SYSTEM\CurrentControlSet\services\Tcpip\Parameters", True)
        subkey.SetValue("IPEnableRouter", 1)

        '' Verify the Service is enabled (Automatic) and started
        Dim obj As ManagementObject
        Dim inParams, outParams As ManagementBaseObject
        Dim Result As Integer
        Dim sc As ServiceController

        obj = New ManagementObject("\\.\root\cimv2:Win32_Service.Name='RemoteAccess'")
        sc = New ServiceController("RemoteAccess")

        'Change the Start Mode to Automatic         
        If obj("StartMode").ToString = "Disabled" Then
            'Get an input parameters object for this method  
            inParams = obj.GetMethodParameters("ChangeStartMode")
            inParams("StartMode") = "Automatic"

            'do it!               
            outParams = obj.InvokeMethod("ChangeStartMode", inParams, Nothing)
            Result = Convert.ToInt32(outParams("returnValue"))
            sc.Start()

            If Result <> 0 Then
                Throw New Exception("ChangeStartMode method error code " & Result)
            End If
        End If

        '' Push persistent static route to routing table
        Dim OpenCMD
        OpenCMD = CreateObject("wscript.shell")
        OpenCMD.run("cmd.exe /C route add -p 10.0.10.0 mask 255.255.255.0 10.181.209.244")

    End Sub


End Class




