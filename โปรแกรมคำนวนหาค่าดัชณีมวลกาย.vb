Public Class โปรแกรมคำนวนหาค่าดัชณีมวลกาย

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim height, weight, bmi As Single
        Height = TextBox1.Text
        Weight = TextBox2.Text
        bmi = (weight) / (height / 100) ^ 2
        Label8.Text = bmi

        If bmi >= 30 Then

            Label9.Text = "อ้วนมาก"
            Label10.Text = "ต้องลดน้ำหนัก ออกกำลังกายอย่างจริงจัง และเข้มงวด"

        ElseIf bmi >= 25 Then

            Label9.Text = "น้ำหนักเกิน"
            Label10.Text = "ต้องลดน้ำหนัก ออกกำลังกายอย่างจริงจัง และเข้มงวด"
        ElseIf bmi >= 23 Then

            Label9.Text = "เริ่มอ้วน"
            Label10.Text = "ต้องลดน้ำหนัก ออกกำลังกายอย่างจริงจัง และเข้มงวด"
        ElseIf bmi >= 18.5 Then

            Label9.Text = "น้ำหนักปกติ"
            Label10.Text = "ต้องเริ่มควบคุมน้ำหนัก ควบคุมอาหาร และออกกำลังกาย"
        Else

            Label9.Text = "ผอมเกินไป"
            Label10.Text = "ต้องเพิ่มน้ำหนัก และรับประทานอาหารที่มีประโยชน์"
        End If

    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ 
                e.Handled = False
            Case 8, 13, 46 'ปุ่ม Backspace = 8,ปุ่ม Enter = 13, ปุ่มDelete = 46
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("สามารถกดได้แค่ตัวเลข")
        End Select
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged


    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ 
                e.Handled = False
            Case 8, 13, 46 'ปุ่ม Backspace = 8,ปุ่ม Enter = 13, ปุ่มDelete = 46
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("สามารถกดได้แค่ตัวเลข")
        End Select
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        Label8.Text = ""
        Label9.Text = ""
        Label10.Text = ""
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Close()
    End Sub
End Class
