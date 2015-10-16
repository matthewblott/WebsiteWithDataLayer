Imports System.Web.Http

Public Class LoginController
  Inherits ApiController

  Public Sub Post(<FromBody()> user As User)

    Dim name = If(user.Name = "matt" And user.Password = "password", "welcome", "login-error")

    HttpContext.Current.Response.Redirect(String.Format("/{0}.html", name))

  End Sub

End Class
