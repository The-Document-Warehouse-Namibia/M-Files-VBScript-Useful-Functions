Private Function GetObjectTypeIDByAlias(objectTypeAlias)
  Dim objectTypeID
  objectTypeID = Vault.ObjectTypeOperations.GetObjectTypeIDByAlias(objectTypeAlias)

  If objectTypeID = -1 Then
    ThrowCustomException NotFoundException("Object Type", objectTypeAlias)
  End If

  GetObjectTypeIDByAlias = objectTypeID
End Function

Private Function GetClassIDByAlias(classAlias)
  Dim classID
  classID = Vault.ClassOperations.GetObjectClassIDByAlias(classAlias)

  If(classID = -1) Then
    ThrowCustomException NotFoundException("Class", classAlias)
  End If

  GetClassIDByAlias = classID
End Function

Private Function GetPropertyDefIDByAlias(propertyAlias)
  Dim propertyDefinitionID
  propertyDefinitionID = Vault.PropertyDefOperations.GetPropertyDefIDByAlias(propertyAlias)

  If propertyDefinitionID = -1 Then
    ThrowCustomException NotFoundException("Property", propertyAlias)
  End If

  GetPropertyDefIDByAlias = propertyDefinitionID
End Function

Private Sub ThrowCustomException(exceptionMessage)
  Err.Raise MFScriptCancel, exceptionMessage
End Sub

Private Function NotFoundException(objType, objAlias)
  NotFoundException = objType & " with alias '" & objAlias & "' does not exist or there is more than one " & LCase(objType) & " with the same alias." & vbCrLf &_
            "Contact your M-Files administrator."
End Function