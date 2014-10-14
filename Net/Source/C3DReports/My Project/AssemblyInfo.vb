Imports System.Resources

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitleAttribute("C3DReport")> 
<Assembly: AssemblyProductAttribute("C3DReport")> 

<Assembly: ComVisibleAttribute(False)> 

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: GuidAttribute("9D6B41AC-BDFE-4e09-9D1D-64362C699B8C")> 

'Add this attribute to avoid scanning all public types.
<Assembly: Autodesk.AutoCAD.Runtime.ExtensionApplication(GetType(ReportApplication))> 
<Assembly: Autodesk.AutoCAD.Runtime.CommandClass(GetType(ReportCommand))> 
'<Assembly: NeutralResourcesLanguageAttribute("en-US")> 
