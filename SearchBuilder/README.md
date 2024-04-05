# SearchBuilder Class

This is a VBScript class `SearchBuilder_` designed for building search conditions using the M-Files API.

## Implementation

Prior to utilizing this class, ensure its inclusion within your script. It can be seamlessly integrated by copying its contents from the MFVBScriptSearchBuilder.vbs file. Placing it at the conclusion of your script or within your functions section is recommended.

This ensures that the class is readily accessible and can be employed efficiently within your VBScript environment.

## Usage

To use this class, follow these steps:

1. **Instantiate the SearchBuilder object:** 
    ```vbscript
    Dim searchBuilder : Set searchBuilder = New SearchBuilder_
    ```

2. **Add search conditions:**

    ```vbscript
    searchBuilder.Deleted deleteStatus
    searchBuilder.ObjType objectTypeID
    searchBuilder.MFClass classID
    searchBuilder.WFState wfStateID
    searchBuilder.PropertyDef propertyDefID, value
    ```

3. **Perform the search:**

    ```vbscript
    Dim searchResults : Set searchResults = searchBuilder.Find()
    ```

## Methods

### Deleted(deleteStatus)

Add a search condition for the deletion status of objects.

- `deleteStatus`: Boolean value indicating whether to search for deleted or non-deleted objects.

### ObjType(objectTypeID)

Add a search condition for the object type ID.

- `objectTypeID`: The ID of the object type to search for.

### MFClass(classID)

Add a search condition for the M-Files class ID.

- `classID`: The ID of the M-Files class to search for.

### WFState(wfStateID)

Add a search condition for the M-Files workflow state ID.

- `wfStateID`: The ID of the M-Files workflow state to search for.
- 
### PropertyDef(propertyDefID, value)

Add a search condition for a custom property definition.

- `propertyDefID`: The ID of the custom property definition.
- `value`: The value to search for.

### Find()

Execute the search and return the search results.

## Example

```vbscript
Dim searchBuilder : Set searchBuilder = New SearchBuilder_

searchBuilder.Deleted False
searchBuilder.ObjType 0
searchBuilder.MFClass 123
searchBuilder.WFState 105
searchBuilder.PropertyDef 1456, "SomeValue"

Dim searchResults : Set searchResults = searchBuilder.Find()
```

