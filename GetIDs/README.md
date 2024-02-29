# M-Files VBScript Utility Functions

This repository contains a set of VBScript utility functions designed to facilitate interaction with M-Files objects. These functions enable easy retrieval of object type IDs, class IDs, and property definition IDs based on their aliases.

## Functions

### GetObjectTypeIDByAlias(objectTypeAlias)

This function retrieves the object type ID corresponding to the provided object type alias.

- **Parameters:**
  - `objectTypeAlias`: Alias of the object type whose ID is to be retrieved.

### GetClassIDByAlias(classAlias)

This function retrieves the class ID corresponding to the provided class alias.

- **Parameters:**
  - `classAlias`: Alias of the class whose ID is to be retrieved.

### GetPropertyDefIDByAlias(propertyAlias)

This function retrieves the property definition ID corresponding to the provided property alias.

- **Parameters:**
  - `propertyAlias`: Alias of the property definition whose ID is to be retrieved.

### ThrowCustomException(exceptionMessage)

This function raises a custom exception with the specified message.

- **Parameters:**
  - `exceptionMessage`: Message to be included in the custom exception.

### NotFoundException(objType, objAlias)

This function generates a message indicating that the specified object type, class, or property does not exist or there are multiple objects with the same alias.

- **Parameters:**
  - `objType`: Type of the object (e.g., "Object Type", "Class", "Property").
  - `objAlias`: Alias of the object for which the exception is raised.

## Usage

Before utilizing these functions, ensure their inclusion within your VBScript environment. You can simply copy and paste the function definitions into your script.

## Note

These functions assume access to an M-Files Vault object (`Vault`). Ensure proper configuration and connection to the M-Files server within your VBScript environment.

For further details on M-Files VBScript development, refer to the [M-Files Developer Portal](https://developer.m-files.com/).
