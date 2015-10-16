Imports System.Web.Http
Imports System.Web.Routing

Public Class AppStart
  Inherits HttpApplication

  Protected Sub Application_Start(sender As Object, e As EventArgs)
    RouteTable.Routes.MapHttpRoute("webApi", "api/{controller}/{action}", New With {.action = RouteParameter.Optional})
  End Sub

End Class