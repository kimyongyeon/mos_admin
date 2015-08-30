<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Debug
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.  
    '코드 편집기를 사용하여 수정하지 마십시오.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button_Close = New System.Windows.Forms.Button
        Me.TextBox_Msg = New System.Windows.Forms.TextBox
        Me.Button_Clear = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Button_Close
        '
        Me.Button_Close.Location = New System.Drawing.Point(328, 234)
        Me.Button_Close.Name = "Button_Close"
        Me.Button_Close.Size = New System.Drawing.Size(96, 30)
        Me.Button_Close.TabIndex = 0
        Me.Button_Close.Text = "닫기"
        Me.Button_Close.UseVisualStyleBackColor = True
        '
        'TextBox_Msg
        '
        Me.TextBox_Msg.Location = New System.Drawing.Point(11, 10)
        Me.TextBox_Msg.Multiline = True
        Me.TextBox_Msg.Name = "TextBox_Msg"
        Me.TextBox_Msg.Size = New System.Drawing.Size(412, 212)
        Me.TextBox_Msg.TabIndex = 1
        '
        'Button_Clear
        '
        Me.Button_Clear.Location = New System.Drawing.Point(234, 234)
        Me.Button_Clear.Name = "Button_Clear"
        Me.Button_Clear.Size = New System.Drawing.Size(87, 32)
        Me.Button_Clear.TabIndex = 2
        Me.Button_Clear.Text = "초기화"
        Me.Button_Clear.UseVisualStyleBackColor = True
        '
        'Form_Debug
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(434, 273)
        Me.Controls.Add(Me.Button_Clear)
        Me.Controls.Add(Me.TextBox_Msg)
        Me.Controls.Add(Me.Button_Close)
        Me.Name = "Form_Debug"
        Me.Text = "Form_Debug"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button_Close As System.Windows.Forms.Button
    Friend WithEvents TextBox_Msg As System.Windows.Forms.TextBox
    Friend WithEvents Button_Clear As System.Windows.Forms.Button
End Class
