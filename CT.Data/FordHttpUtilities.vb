
Imports System.Net

Public Class FordHttpUtilities

    Public Shared Function PostToWebPage(ByVal endPoint As String, ByVal wslCookie As String, ByVal httpBody As String) As String
        Return CallWebMethod(endPoint, wslCookie, HttpMethodConstants.HTTP_POST, httpBody)
    End Function

    Public Shared Function GetWebPage(ByVal endPoint As String, ByVal wslCookie As String) As String
        Return CallWebMethod(endPoint, wslCookie, HttpMethodConstants.HTTP_GET, "")
    End Function

    Private Shared Function CallWebMethod(
            ByVal endPoint As String,
            ByVal wslCookie As String,
            ByVal httpMethod As String,
            ByVal httpBody As String) As String

        Const SXH_OPTION_SELECT_CLIENT_SSL_CERT = 3
        Dim objXmlHttp 'As MSXML2.ServerXMLHTTP = Nothing
        Dim _objXmlHttp As HttpWebRequest()
        Try
            Dim strDigitalCert As String
            strDigitalCert = "LOCAL_MACHINE/My/Certificates"
            Dim intStatus

            'objXmlHttp = New MSXML2.ServerXMLHTTP

            'passing in the digital certificate required for SSL
            objXmlHttp.setOption(SXH_OPTION_SELECT_CLIENT_SSL_CERT, strDigitalCert)

            ' ignore all certificate error
            ' should not be used if in production assuming all certs should be valid
            ' and if not, there should be error.
            ' The following two line should be commented out when in Production.
            'Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
            'objXmlHttp.setOption(2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)

            ' This uses POST, you can use GET too
            objXmlHttp.open(httpMethod, endPoint, False)
            If Trim(wslCookie) <> "" Then
                'passing in the WSL cookie value
                objXmlHttp.setRequestHeader("COOKIE", wslCookie)
            End If

            If httpMethod = HttpMethodConstants.HTTP_GET Then
                objXmlHttp.send()
            Else
                objXmlHttp.send(httpBody)
            End If

            intStatus = objXmlHttp.status
            If (intStatus <> 200) Then
                System.Diagnostics.Debug.Print("The status for the HTTP request call returned is: " & intStatus)
                Throw New Exception("The status for the HTTP request call returned is: " & intStatus)
            End If

            Return objXmlHttp.responseText
        Catch ex As Exception
            Throw
        Finally
            If objXmlHttp IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objXmlHttp)
                objXmlHttp = Nothing
            End If
        End Try
    End Function

    Private Structure HttpMethodConstants
        Public Const HTTP_POST As String = "POST"
        Public Const HTTP_GET As String = "GET"
    End Structure


End Class
