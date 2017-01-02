Option Explicit On
Option Strict On
Imports System.Transactions

Public Class Form1
    Dim db As New dbNorthwindDataContext()

    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        Dim od1 As New Order_Detail()
        od1.ProductID = 1
        od1.Quantity = 10
        od1.UnitPrice = 100
        od1.Discount = 0

        Dim od2 As New Order_Detail()
        od2.ProductID = 2
        od2.Quantity = 20
        od2.UnitPrice = 100
        od2.Discount = 0

        Dim o As New Order()
        o.CustomerID = "BERGS"
        o.OrderDate = Date.Now
        o.EmployeeID = 5
        o.RequiredDate = Date.Now.AddDays(3)
        o.Order_Details.Add(od1)
        o.Order_Details.Add(od2)

        Using ts As New TransactionScope()
            db.Orders.InsertOnSubmit(o)
            db.SubmitChanges()
            ts.Complete()
        End Using
        MessageBox.Show("เพิ่มข้อมูลเรียบร้อยแล้ว!!", "ผลการทำงาน")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If txtOrderID.Text.Trim() = "" Then Exit Sub
        Try
            Dim ods = From od In db.Order_Details
                      Where od.Order.OrderID = CDbl(txtOrderID.Text)
                      Select od

            For Each od In ods
                db.Order_Details.DeleteAllOnSubmit(ods)
                'db.Order_Details.DeleteOnSubmit(od)
            Next

            Dim os = (From o In db.Orders
                      Where o.OrderID = CDbl(txtOrderID.Text)
                      Select o).FirstOrDefault()

            Using ts As New TransactionScope()
                db.Orders.DeleteOnSubmit(os)
                db.SubmitChanges()
                ts.Complete()
            End Using
            MessageBox.Show("ลบข้อมูลเรียบร้อยแล้ว !!!", "ผลการทำงาน")

        Catch ex As Exception
            MessageBox.Show("คุณป้อนรหัสสั่งซื้อสินค้า ไม่ถูกต้อง!!", "แจ้งเตือน")
            txtOrderID.Focus()
            txtOrderID.SelectAll()
        End Try
    End Sub
End Class
