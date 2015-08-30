<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Result
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
        Me.Label_Message = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label_ExamTitle = New System.Windows.Forms.Label
        Me.Label_PassScore = New System.Windows.Forms.Label
        Me.Label_TotalScore = New System.Windows.Forms.Label
        Me.Label_Result = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label_Message
        '
        Me.Label_Message.AutoSize = True
        Me.Label_Message.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label_Message.Location = New System.Drawing.Point(24, 17)
        Me.Label_Message.Name = "Label_Message"
        Me.Label_Message.Size = New System.Drawing.Size(207, 12)
        Me.Label_Message.TabIndex = 0
        Me.Label_Message.Text = "축하합니다! 시험에 합격했습니다!"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.Location = New System.Drawing.Point(24, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 12)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "시험 이름:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.Location = New System.Drawing.Point(24, 70)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(113, 12)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "합격에 필요한 점수:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.Location = New System.Drawing.Point(24, 92)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 12)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "귀하의 점수:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("굴림", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label4.Location = New System.Drawing.Point(24, 117)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(33, 12)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "결과:"
        '
        'Label_ExamTitle
        '
        Me.Label_ExamTitle.AutoSize = True
        Me.Label_ExamTitle.Location = New System.Drawing.Point(241, 48)
        Me.Label_ExamTitle.Name = "Label_ExamTitle"
        Me.Label_ExamTitle.Size = New System.Drawing.Size(161, 12)
        Me.Label_ExamTitle.TabIndex = 5
        Me.Label_ExamTitle.Text = "Microsoft Office 2003 Expert"
        '
        'Label_PassScore
        '
        Me.Label_PassScore.AutoSize = True
        Me.Label_PassScore.Location = New System.Drawing.Point(241, 70)
        Me.Label_PassScore.Name = "Label_PassScore"
        Me.Label_PassScore.Size = New System.Drawing.Size(23, 12)
        Me.Label_PassScore.TabIndex = 6
        Me.Label_PassScore.Text = "650"
        '
        'Label_TotalScore
        '
        Me.Label_TotalScore.AutoSize = True
        Me.Label_TotalScore.Location = New System.Drawing.Point(241, 92)
        Me.Label_TotalScore.Name = "Label_TotalScore"
        Me.Label_TotalScore.Size = New System.Drawing.Size(23, 12)
        Me.Label_TotalScore.TabIndex = 7
        Me.Label_TotalScore.Text = "900"
        '
        'Label_Result
        '
        Me.Label_Result.AutoSize = True
        Me.Label_Result.Location = New System.Drawing.Point(241, 117)
        Me.Label_Result.Name = "Label_Result"
        Me.Label_Result.Size = New System.Drawing.Size(29, 12)
        Me.Label_Result.TabIndex = 8
        Me.Label_Result.Text = "합격"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(408, 172)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(213, 28)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "성적표 인쇄(P)"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form_Result
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(631, 214)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label_Result)
        Me.Controls.Add(Me.Label_TotalScore)
        Me.Controls.Add(Me.Label_PassScore)
        Me.Controls.Add(Me.Label_ExamTitle)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label_Message)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form_Result"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "MOS Simulation Program V1.2 - 시험 점수"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label_Message As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label_ExamTitle As System.Windows.Forms.Label
    Friend WithEvents Label_PassScore As System.Windows.Forms.Label
    Friend WithEvents Label_TotalScore As System.Windows.Forms.Label
    Friend WithEvents Label_Result As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
