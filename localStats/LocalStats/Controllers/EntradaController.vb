Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Entity
Imports System.Linq
Imports System.Net
Imports System.Web
Imports System.Web.Mvc
Imports LocalStats

Namespace Controllers
    Public Class EntradaController
        Inherits System.Web.Mvc.Controller

        Private db As New bdContext

        ' GET: Entrada
        Function Index() As ActionResult
            Return View(db.Movies.ToList())
        End Function

        ' GET: Entrada/Details/5
        Function Details(ByVal id As Integer?) As ActionResult
            If IsNothing(id) Then
                Return New HttpStatusCodeResult(HttpStatusCode.BadRequest)
            End If
            Dim input As input = db.Movies.Find(id)
            If IsNothing(input) Then
                Return HttpNotFound()
            End If
            Return View(input)
        End Function

        ' GET: Entrada/Create
        Function Create() As ActionResult
            Return View()
        End Function

        ' POST: Entrada/Create
        'To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        'more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        <HttpPost()>
        <ValidateAntiForgeryToken()>
        Function Create(<Bind(Include:="ID,local,utente,qty")> ByVal input As input) As ActionResult
            If ModelState.IsValid Then
                db.Movies.Add(input)
                db.SaveChanges()
                Return RedirectToAction("Create")
            End If
            Return View(input)
        End Function

        ' GET: Entrada/Edit/5
        Function Edit(ByVal id As Integer?) As ActionResult
            If IsNothing(id) Then
                Return New HttpStatusCodeResult(HttpStatusCode.BadRequest)
            End If
            Dim input As input = db.Movies.Find(id)
            If IsNothing(input) Then
                Return HttpNotFound()
            End If
            Return View(input)
        End Function

        ' POST: Entrada/Edit/5
        'To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        'more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        <HttpPost()>
        <ValidateAntiForgeryToken()>
        Function Edit(<Bind(Include:="ID,local,utente,qty")> ByVal input As input) As ActionResult
            If ModelState.IsValid Then
                db.Entry(input).State = EntityState.Modified
                db.SaveChanges()
                Return RedirectToAction("Index")
            End If
            Return View(input)
        End Function

        ' GET: Entrada/Delete/5
        Function Delete(ByVal id As Integer?) As ActionResult
            If IsNothing(id) Then
                Return New HttpStatusCodeResult(HttpStatusCode.BadRequest)
            End If
            Dim input As input = db.Movies.Find(id)
            If IsNothing(input) Then
                Return HttpNotFound()
            End If
            Return View(input)
        End Function

        ' POST: Entrada/Delete/5
        <HttpPost()>
        <ActionName("Delete")>
        <ValidateAntiForgeryToken()>
        Function DeleteConfirmed(ByVal id As Integer) As ActionResult
            Dim input As input = db.Movies.Find(id)
            db.Movies.Remove(input)
            db.SaveChanges()
            Return RedirectToAction("Index")
        End Function

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If (disposing) Then
                db.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub
    End Class
End Namespace
