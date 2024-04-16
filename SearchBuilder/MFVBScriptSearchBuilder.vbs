Class SearchBuilder_
	Private searchConditions_
	Private searchCondition_

	Private Sub Class_Initialize()
		Set searchConditions_ = CreateObject("MFilesAPI.SearchConditions")
	End Sub
	
	Public Sub Deleted(deleteStatus)
		Set searchCondition_ = CreateObject("MFilesAPI.SearchCondition")
		
		searchCondition_.Expression.DataStatusValueType = MFStatusTypeDeleted
		searchCondition_.ConditionType = MFConditionTypeEqual
		searchCondition_.TypedValue.SetValue MFDatatypeBoolean, deleteStatus
		
		searchConditions_.Add -1, searchCondition_											
	End Sub

	Public Sub ObjType(objectTypeID)
		Set searchCondition_ = CreateObject("MFilesAPI.SearchCondition")
		
		searchCondition_.Expression.DataStatusValueType = MFStatusTypeObjectTypeID
		searchCondition_.ConditionType = MFConditionTypeEqual
		searchCondition_.TypedValue.SetValue MFDatatypeLookup, objectTypeID

		searchConditions_.Add -1, searchCondition_
	End Sub
	
	Public Sub MFClass(classID)
		Set searchCondition_ = CreateObject("MFilesAPI.SearchCondition")

		searchCondition_.Expression.DataPropertyValuePropertyDef = MFBuiltInPropertyDefClass
		searchCondition_.ConditionType = MFConditionTypeEqual
		searchCondition_.TypedValue.SetValue MFDatatypeLookup, classID

		searchConditions_.Add -1, searchCondition_
	End Sub
	
	Public Sub WFState(wfStateID)
		Set searchCondition_ = CreateObject("MFilesAPI.SearchCondition")

		searchCondition_.Expression.DataPropertyValuePropertyDef = MFBuiltInPropertyDefState
		searchCondition_.ConditionType = MFConditionTypeEqual
		searchCondition_.TypedValue.SetValue MFDatatypeLookup, wfStateID

		searchConditions_.Add -1, searchCondition_
	End Sub
	
	Public Sub PropertyDef(propertyDefID, propertyDefDataType, value)
		Set searchCondition_ = CreateObject("MFilesAPI.SearchCondition")

		searchCondition_.Expression.DataPropertyValuePropertyDef = propertyDefID
		searchCondition_.ConditionType = MFConditionTypeEqual
		searchCondition_.TypedValue.SetValue propertyDefDataType, value

		searchConditions_.Add -1, searchCondition_
	End Sub

	Public Function Find()
		Set Find = Vault.ObjectSearchOperations.SearchForObjectsByConditions(searchConditions_, MFSearchFlagNone, False)
	End Function
End Class
