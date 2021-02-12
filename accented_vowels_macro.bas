Attribute VB_Name = "NewMacros"
''' ALT + A
Sub accented_a()
    Word.Application.Selection.Text = "à"
    Application.Selection.EndOf
End Sub

''' ALT + CTRL + e
Sub accented_e_up()
    Word.Application.Selection.Text = "é"
    Application.Selection.EndOf
End Sub

''' ALT + e
Sub accented_e_down()
    Word.Application.Selection.Text = "è"
    Application.Selection.EndOf
End Sub

''' ALT + i
Sub accented_i()
    Word.Application.Selection.Text = "ì"
    Application.Selection.EndOf
End Sub

''' ALT + o
Sub accented_o_down()
    Word.Application.Selection.Text = "ò"
    Application.Selection.EndOf
End Sub

''' ALT + CTRL + o
Sub accented_o_up()
    Word.Application.Selection.Text = "ó"
    Application.Selection.EndOf
End Sub

''' ALT + u
Sub accented_u()
    Word.Application.Selection.Text = "ù"
    Application.Selection.EndOf
End Sub

''' ALT + SHIFT + 4
Sub euro()
    Word.Application.Selection.Text = "€"
    Application.Selection.EndOf
End Sub










