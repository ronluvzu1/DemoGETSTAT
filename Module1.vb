Module Module1

    Sub Main()

        Console.Write("----Remote Host------" & vbCrLf)
        Console.Write("Remote Machine name: ")
        Dim strHost = Console.ReadLine()
        Console.Write("Username: ")
        Dim userName = Console.ReadLine()
        Console.Write("Password: ")
        Dim passWord = Console.ReadLine()
        'Dim strComputer = "."
        Dim objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
        Try
            Dim objSWbemServices = objSWbemLocator.ConnectServer(strHost, "root\cimv2", userName, passWord)

            Dim colSWbemObjectSet = objSWbemServices.ExecQuery("SELECT * FROM Win32_ComputerSystem")

            For Each objSWbemObject In colSWbemObjectSet
                Console.WriteLine("Name: " & objSWbemObject.Name & vbCrLf &
                              "Description: " & objSWbemObject.Description & vbCrLf &
                              "DNSHostName: {0} ", objSWbemObject.DNSHostName & vbCrLf &
                              "Domain: " & objSWbemObject.Domain & vbCrLf &
                              "Model: " & objSWbemObject.Model & vbCrLf &
                              "NumberOfProcessors:   " & objSWbemObject.NumberOfProcessors & vbCrLf &
                              "TotalPhysicalMemory: " & objSWbemObject.TotalPhysicalMemory & vbCrLf &
                              "AdminPasswordStatus: " & objSWbemObject.AdminPasswordStatus & vbCrLf &
                              "Manufacturer: " & objSWbemObject.Manufacturer & vbCrLf)
            Next
        Catch ex As Exception
            Console.WriteLine("Remote Feature Error" & vbCrLf &
                              "*HResult*: {0} *Source*: {1} *ex Message* {2}  ", ex.HResult, ex.Source, ex.Message)
        End Try



        Console.Write("----Local Host------" & vbCrLf)

        Dim localstrHost = ""


        'Dim strComputer = "."
        Dim localobjSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
        Try
            Dim localobjSWbemServices = localobjSWbemLocator.ConnectServer(localstrHost, "root\cimv2")
            Dim localcolSWbemObjectSet = localobjSWbemServices.ExecQuery("SELECT * FROM Win32_ComputerSystem")

            For Each localobjSWbemObject In localcolSWbemObjectSet
                Console.WriteLine("Name: " & localobjSWbemObject.Name & vbCrLf &
                              "Description: " & localobjSWbemObject.Description & vbCrLf &
                              "DNSHostName: {0} ", localobjSWbemObject.DNSHostName & vbCrLf &
                              "Domain: " & localobjSWbemObject.Domain & vbCrLf &
                              "Model: " & localobjSWbemObject.Model & vbCrLf &
                              "NumberOfProcessors:   " & localobjSWbemObject.NumberOfProcessors & vbCrLf &
                              "TotalPhysicalMemory: " & localobjSWbemObject.TotalPhysicalMemory & vbCrLf &
                              "AdminPasswordStatus: " & localobjSWbemObject.AdminPasswordStatus & vbCrLf &
                              "Manufacturer: " & localobjSWbemObject.Manufacturer)


                Console.ReadLine()
            Next
        Catch ex As Exception
            Console.WriteLine("Remote Feature Error" & vbCrLf &
                              "HResult: {0} Source: {1} ex Message {2}  ", ex.HResult, ex.Source, ex.Message)
        End Try

    End Sub
End Module


