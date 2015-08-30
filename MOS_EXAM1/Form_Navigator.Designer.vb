<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Navigator
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
        Me.components = New System.ComponentModel.Container
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.RichTextBox_Question = New System.Windows.Forms.RichTextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label_Timer = New System.Windows.Forms.Label
        Me.Label_Seq = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Button_Next = New System.Windows.Forms.Button
        Me.Button_Retry = New System.Windows.Forms.Button
        Me.Button_Skip = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.Button_End = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RichTextBox_Question)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(835, 108)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'RichTextBox_Question
        '
        Me.RichTextBox_Question.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox_Question.Location = New System.Drawing.Point(3, 17)
        Me.RichTextBox_Question.Name = "RichTextBox_Question"
        Me.RichTextBox_Question.Size = New System.Drawing.Size(829, 88)
        Me.RichTextBox_Question.TabIndex = 9
        Me.RichTextBox_Question.Text = ""
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label_Timer)
        Me.GroupBox2.Controls.Add(Me.Label_Seq)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Left
        Me.GroupBox2.Location = New System.Drawing.Point(0, 108)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox2.Size = New System.Drawing.Size(203, 42)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'Label_Timer
        '
        Me.Label_Timer.AutoSize = True
        Me.Label_Timer.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label_Timer.ForeColor = System.Drawing.Color.White
        Me.Label_Timer.Location = New System.Drawing.Point(110, 20)
        Me.Label_Timer.Name = "Label_Timer"
        Me.Label_Timer.Size = New System.Drawing.Size(57, 12)
        Me.Label_Timer.TabIndex = 1
        Me.Label_Timer.Text = "00:49:33"
        '
        'Label_Seq
        '
        Me.Label_Seq.AutoSize = True
        Me.Label_Seq.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label_Seq.ForeColor = System.Drawing.Color.White
        Me.Label_Seq.Location = New System.Drawing.Point(9, 20)
        Me.Label_Seq.Name = "Label_Seq"
        Me.Label_Seq.Size = New System.Drawing.Size(49, 12)
        Me.Label_Seq.TabIndex = 0
        Me.Label_Seq.Text = "19 중 1"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Button_Next)
        Me.GroupBox3.Controls.Add(Me.Button_Retry)
        Me.GroupBox3.Controls.Add(Me.Button_Skip)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Right
        Me.GroupBox3.Location = New System.Drawing.Point(562, 108)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox3.Size = New System.Drawing.Size(273, 42)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        '
        'Button_Next
        '
        Me.Button_Next.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Button_Next.Location = New System.Drawing.Point(191, 12)
        Me.Button_Next.Name = "Button_Next"
        Me.Button_Next.Size = New System.Drawing.Size(79, 26)
        Me.Button_Next.TabIndex = 2
        Me.Button_Next.Text = "다음"
        Me.Button_Next.UseVisualStyleBackColor = False
        '
        'Button_Retry
        '
        Me.Button_Retry.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Button_Retry.Location = New System.Drawing.Point(110, 12)
        Me.Button_Retry.Name = "Button_Retry"
        Me.Button_Retry.Size = New System.Drawing.Size(79, 26)
        Me.Button_Retry.TabIndex = 1
        Me.Button_Retry.Text = "다시풀기"
        Me.Button_Retry.UseVisualStyleBackColor = False
        '
        'Button_Skip
        '
        Me.Button_Skip.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Button_Skip.Location = New System.Drawing.Point(25, 12)
        Me.Button_Skip.Name = "Button_Skip"
        Me.Button_Skip.Size = New System.Drawing.Size(79, 26)
        Me.Button_Skip.TabIndex = 0
        Me.Button_Skip.Text = "건너뛰기"
        Me.Button_Skip.UseVisualStyleBackColor = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Button_End)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox4.Location = New System.Drawing.Point(203, 108)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(0)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox4.Size = New System.Drawing.Size(359, 42)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        '
        'Button_End
        '
        Me.Button_End.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.Button_End.Location = New System.Drawing.Point(256, 12)
        Me.Button_End.Name = "Button_End"
        Me.Button_End.Size = New System.Drawing.Size(97, 25)
        Me.Button_End.TabIndex = 17
        Me.Button_End.Text = "종료"
        Me.Button_End.UseVisualStyleBackColor = False
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'Form_Navigator
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(835, 150)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form_Navigator"
        Me.Text = "Form_Navigator"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label_Timer As System.Windows.Forms.Label
    Friend WithEvents Label_Seq As System.Windows.Forms.Label
    Friend WithEvents Button_Next As System.Windows.Forms.Button
    Friend WithEvents Button_Retry As System.Windows.Forms.Button
    Friend WithEvents Button_Skip As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Button_End As System.Windows.Forms.Button
    Friend WithEvents RichTextBox_Question As System.Windows.Forms.RichTextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
