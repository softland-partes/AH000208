'Seccion de declaracion de variables
Dim m_bHayErrores, m_sIsmulti, m_sCodemp, m_sCodaplicacionGA,  m_lNrolot, oReport, m_iRecuperaDatos
'--------------------------------------------------------------------------------------------------------------------------
Function On_Initialization(sErrorMessage)
  sErrorMessage = InicializoVariablesGlobales
  m_lNrolot = TMinstance.Table.rows(1).fields("GRRMDH_NROLOT").Value
  TMinstance.Table.rows(1).fields("USR_GRRMDH_SELTOD").Value = "S"
  m_iRecuperaDatos = 0

End Function
'--------------------------------------------------------------------------------------------------------------------------
Function Previous_SaveAplication(sErrorMessage)
End Function
'--------------------------------------------------------------------------------------------------------------------------
Function On_SaveAplication(sErrorMessage)
End Function
'--------------------------------------------------------------------------------------------------------------------------
Function Pos_SaveAplication(sErrorMessage)
    sSql =  "UPDATE CVMCTI SET "
    sSql = sSql & "CVMCTI_PRECIO = USR_AH0193_NEWPRE "
    sSql = sSql & "FROM USR_AH0193 "
    sSql = sSql & "INNER JOIN CVMCTI ON "
    sSql = sSql & "CVMCTI_CODCON = USR_AH0193_CODCON AND "
    sSql = sSql & "CVMCTI_NROCON = USR_AH0193_NROCON AND "
    sSql = sSql & "CVMCTI_NROEXT= USR_AH0193_NROEXT AND "
    sSql = sSql & "CVMCTI_NROITM = USR_AH0193_ITMCON  "
    sSql = sSql & "WHERE "
    sSql = sSql & "USR_AH0193_NROLOT ="& m_lNrolot &" AND "
    sSql = sSql & "USR_AH0193_NEWPRE > 0 "
    sSql = sSql & " and CVMCTI_CODEMP = '"&m_sCodemp&"'"
    sSql = sSql & " and USR_AH0193_SELTOD ='S'"
    Set oRd = TMinstance.openresultset(CStr(sSql))

End Function
'--------------------------------------------------------------------------------------------------------------------------
Function USR_GRRMDH_TIPPRODES_On_Change(oField, sErrorMessage)
    TMinstance.Table.rows(1).fields("USR_GRRMDH_TIPPROHAS").Value = oField.value
End Function
Function USR_GRRMDH_ARTCODDES_On_Change(oField, sErrorMessage)
    TMinstance.Table.rows(1).fields("USR_GRRMDH_ARTCODHAS").Value = oField.value
End Function
Function USR_GRRMDH_RECUPE_On_Change(oField, sErrorMessage)
  Dim sSql, oRd, oColumn

  with TMinstance.Table.rows(1)
    If .fields("USR_GRRMDH_RECUPE").Value = "S" then

      sSql = "SELECT CVMCTI_CODCON CODCON, CVMCTI_NROCON NROCON, CVMCTI_NROEXT NROEXT, CVMCTI_NROITM ITMCON,  "
      sSql = sSql & "CVMCTH_NROCTA NROCTA, CVMCTI_TIPPRO TIPPRO, CVMCTI_ARTCOD ARTCOD, CVMCTI_PRECIO PRECIO, "
      sSql = sSql & "CVMCTI_CANTID CANTID, USR_CVMCTI_DESIRT DESIRT, USR_CVMCTI_HASIRT HASIRT "
      sSql = sSql & "FROM CVMCTI INNER JOIN CVMCTH ON CVMCTH_CODEMP = CVMCTI_CODEMP AND CVMCTH_NROCON = CVMCTI_NROCON "
      sSql = sSql & "AND CVMCTH_CODCON = CVMCTI_CODCON AND CVMCTH_NROEXT = CVMCTI_NROEXT "
      sSql = sSql & "WHERE NOT EXISTS(SELECT * FROM VTMCLH WHERE VTMCLH_INHIBE = 'S' AND VTMCLH_NROCTA = CVMCTH_NROCTA) AND "
      sSql = sSql & "USR_CVMCTI_CANSER<>'S' AND "'
      if .fields("USR_GRRMDH_TIPPRODES").Value <>"" AND .fields("USR_GRRMDH_TIPPROHAS").Value<>"" THEN
        sSql = sSql & "CVMCTI_TIPPRO BETWEEN '"& .fields("USR_GRRMDH_TIPPRODES").Value &"' AND '"& .fields("USR_GRRMDH_TIPPROHAS").Value &"' AND "
        if  .fields("USR_GRRMDH_ARTCODDES").Value<>"" AND .fields("USR_GRRMDH_ARTCODHAS").Value<>"" THEN
          sSql = sSql & " CVMCTI_ARTCOD BETWEEN '"& .fields("USR_GRRMDH_ARTCODDES").Value &"' AND '"& .fields("USR_GRRMDH_ARTCODHAS").Value &"' AND "
        End if
      End if
      sSql = sSql & "USR_CVMCTI_DESIRT <= '"& .fields("USR_GRRMDH_FCHHAS").Value  &"'"
	     sSql = sSql & " and  CVMCTI_codemp = '"&m_sCodemp&"'"

      Set oRd = TMinstance.openresultset(CStr(sSql))

      m_iRecuperaDatos = oRd.RowCount

      Do While Not oRd.EOF
          With TMinstance.Table.rows(1).tables("USR_AH0193").rows
              For Each Field In .Add.fields
                  For Each oColumn In oRd.rdocolumns
                    sCampo = Replace(Field.Name,"USR_AH0193_","")
                    if sCampo = "SELTOD" then
                      Field.Enabled = true
                      Field.Value = "S"
                    End if
                      If oColumn.Name = sCampo Then
                        if Field.Name = "USR_AH0193_NEWPRE" then
                          Field.Enabled = true
                        Else
                          Field.Enabled = false
                        End if
                          Field.description = recupDescrp(Field.Name)
                          Field.Value = ResuelvoSegunType(oRd(oColumn.Name))
                          Exit For
                      End If
                  Next
              Next
          End With
          oRd.movenext
      Loop
      oRd.Close
      Set oRd = Nothing
    else
      TMinstance.Table.rows(1).tables("USR_AH0193").rows.clear
    End if
    TMinstance.Table.rows(1).tables("USR_AH0193").Rows.MaxRows = m_iRecuperaDatos
  End with
End Function
Function USR_GRRMDH_SELTOD_On_Change(oField, sErrorMessage)
  	If TMinstance.Table.rows(1).fields("USR_GRRMDH_SELTOD").Value = "S" then
		With TMinstance.Table.rows(1).tables("USR_AH0193")
			For each Field in .Rows
			 	Field.Fields("USR_AH0193_SELTOD").Value = "S"
			Next
		End With
	else
		With TMinstance.Table.rows(1).tables("USR_AH0193")
			For each Field in .Rows
				Field.Fields("USR_AH0193_SELTOD").Value = "N"
			Next
		End with
	End If
End Function
Function USR_GRRMDH_PORCEN_On_Change(oField, sErrorMessage)
  if m_iRecuperaDatos = 0 Then
    sErrorMessage = "Debe recuperar datos de contratos antes de indicar el porcentaje de aumento"
    TMinstance.Table.rows(1).fields("USR_GRRMDH_PORCEN").Value = 0
  else
    With TMinstance.Table.rows(1).tables("USR_AH0193")
      For Each oRow In .Rows
        With oRow
          .fields("USR_AH0193_NEWPRE").Value = .fields("USR_AH0193_PRECIO").Value * (1+TMinstance.Table.rows(1).fields("USR_GRRMDH_PORCEN").Value/100)
        End With
      Next
    End With
  End if
End Function
Function ResuelvoSegunType(RdField)
    Select Case RdField.Type
        Case 11
            ResuelvoSegunType = Year(RdField.Value) * 10000 + Month(RdField.Value) * 100 + Day(RdField.Value)
        Case 2
            ResuelvoSegunType = "0" & RdField.Value
        Case Else
            ResuelvoSegunType = RdField.Value
    End Select

End Function
Function InicializoVariablesGlobales()

    Dim sSql, oRd
    Dim sKey, sValue

    On Error Resume Next
    With TMinstance.Table.rows(1)
        m_sCodemp = .fields("USR_GRRMDH_CODEMP").Value
        m_sCodaplicacionGA = .fields("GRRMDH_CODIGO").Value
    End With
    If Err.Number <> 0 Then
        InicializoVariablesGlobales = "El campo USR_GRRMDH_CODEMP no existe, solicite asistencia."
        Exit Function
    End If
    On Error GoTo 0

    sSql = " SELECT * FROM cwSGCore.DBO.CWOMCOMPANIES WHERE NAME = '" & m_sCodemp & "' "
    Set oRd = TMinstance.openresultset(CStr(sSql))
    m_sIsmulti = oRd("ISMULTI").Value

    oRd.Close
    Set oRd = Nothing
  End Function
Sub grabarLog_Archivo(pDato)
      Dim strArchivo, archivo, fso
      Dim ParaEscritura, ParaAnexar

      Set fso = CreateObject("Scripting.FileSystemObject")
      strArchivo = "C:\log\File_" + Replace(Replace(CStr(Date), "/", "-"), ":", ".") + ".log"
      If FileExists(strArchivo) Then
          ParaAnexar = 8
          Set archivo = fso.OpenTextFile(strArchivo, ParaAnexar, False)
      Else
          ParaEscritura = 2
          Set archivo = fso.CreateTextFile(strArchivo)
      End If

      archivo.Write (CStr(Now) + " - " + pDato + vbCRLF)
      archivo.Close
End Sub
Function FileExists(fileName)
    ' aqui On Error Resume Next
    Dim objFso
    Set objFso = CreateObject("Scripting.FileSystemObject")
    If objFso.FileExists(fileName) Then
        FileExists = True
    Else
        FileExists = False
    End If
    If Err > 0 Then Err.Clear
    Set objFso = Nothing
End Function
Function recupDescrp(campo)
  Dim sSql, oRd
    sSql = "SELECT Distinct LongCaption_ARG DESCRP FROM CWTMFIELDS WHERE FIELDNAME = '" & campo & "'"
    Set oRd = TMinstance.openresultset(CStr(sSql))
    if Not oRd.EOF then
      RecupDescrp = oRd("DESCRP").Value
    End if
End Function
