<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Login
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txt_emp_id = New System.Windows.Forms.TextBox()
        Me.txt_emp_pwd = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btn_login = New System.Windows.Forms.Button()
        Me.btn_close = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txt_emp_id
        '
        Me.txt_emp_id.Location = New System.Drawing.Point(92, 59)
        Me.txt_emp_id.Name = "txt_emp_id"
        Me.txt_emp_id.Size = New System.Drawing.Size(100, 22)
        Me.txt_emp_id.TabIndex = 0
        '
        'txt_emp_pwd
        '
        Me.txt_emp_pwd.Location = New System.Drawing.Point(92, 119)
        Me.txt_emp_pwd.Name = "txt_emp_pwd"
        Me.txt_emp_pwd.Size = New System.Drawing.Size(100, 22)
        Me.txt_emp_pwd.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(100, 42)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(29, 12)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "帳號"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(100, 104)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 12)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "密碼"
        '
        'btn_login
        '
        Me.btn_login.Location = New System.Drawing.Point(67, 157)
        Me.btn_login.Name = "btn_login"
        Me.btn_login.Size = New System.Drawing.Size(75, 23)
        Me.btn_login.TabIndex = 4
        Me.btn_login.Text = "登入"
        Me.btn_login.UseVisualStyleBackColor = True
        '
        'btn_close
        '
        Me.btn_close.Location = New System.Drawing.Point(148, 157)
        Me.btn_close.Name = "btn_close"
        Me.btn_close.Size = New System.Drawing.Size(75, 23)
        Me.btn_close.TabIndex = 5
        Me.btn_close.Text = "關閉"
        Me.btn_close.UseVisualStyleBackColor = True
        '
        'Login
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Controls.Add(Me.btn_close)
        Me.Controls.Add(Me.btn_login)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txt_emp_pwd)
        Me.Controls.Add(Me.txt_emp_id)
        Me.Name = "Login"
        Me.Text = "Login"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txt_emp_id As TextBox
    Friend WithEvents txt_emp_pwd As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents btn_login As Button
    Friend WithEvents btn_close As Button
End Class
