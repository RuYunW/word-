Sub �����޸�()
'
' �����޸� ��
'
'
Dim R_Character As Range

Dim FontSize(5)
'������5��ֵ֮����в���
FontSize(1) = "16"
FontSize(2) = "16.2"
FontSize(3) = "16.5"
FontSize(4) = "17"
FontSize(5) = "17.2"

Dim ParagraphSpace(5)
'�м��������ֵ�о��ȷֲ�
ParagraphSpace(1) = "12"
ParagraphSpace(2) = "13"
ParagraphSpace(3) = "17"
ParagraphSpace(4) = "9"
ParagraphSpace(5) = "12"

For Each R_Character In ActiveDocument.Characters
    VBA.Randomize
    R_Character.Font.Size = FontSize(Int(VBA.Rnd * 5) + 1)
    R_Character.Font.Position = Int(VBA.Rnd * 3) + 1
    R_Character.Font.Spacing = 0

Next
Application.ScreenUpdating = True

For Each Cur_Paragraph In ActiveDocument.Paragraphs
    Cur_Paragraph.LineSpacing = ParagraphSpace(Int(VBA.Rnd * 5) + 1)
    
Next
    Application.ScreenUpdating = True
    


End Sub