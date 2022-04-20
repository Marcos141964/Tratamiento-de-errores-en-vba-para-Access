# Tratamiento-de-errores-en-vba-para-Access
 '=======================================
  'Autor: Marcos José López de Dios
  '=======================================
Exit_1:
          DoCmd.Hourglass False
          DoCmd.Echo True
          Exit Sub
1 a:
          DoCmd.Hourglass False
          DoCmd.Echo True
          strMsg = "Erro # " & Str(Err.Number) _
              & vbNewLine & "Descripción: " & Err.Description _
              & vbNewLine & vbNewLine & "Póngase en contacto con el administrador del sistema."
          MsgBox strMsg, vbExclamation, "Atención"
          Resume Exit_1
