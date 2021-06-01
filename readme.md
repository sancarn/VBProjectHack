# Simple VBProject hack

This project contains a simple mechanism which always ensures VBProject is available.

## Usage

The project exposes a class. When the class initialises it will set VBProject extensibility mode to true. When the class de-initialises it will reset the VBProject extensibility.

```vb
Sub doSomething()
    With new VBProjectHack
        'Do work...
        Debug.Print ThisWorkbook.VBProject.Name
    End with
End Sub
```

