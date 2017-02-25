## Getting the Microsoft.Office.Interop.Visio.dll assembly from the GAC

Launch CMD

Run this command

```
SUBST X: C:\Windows\assembly
```

CD to the location of the Assembly

```
CD X:\GAC_MSIL\Microsoft.Office.Interop.Visio\14.0.0.0__71e9bce111e9429c
```

Copy the **Microsoft.Office.Interop.dll** out of the folder

```
COPY X:\GAC_MSIL\Microsoft.Office.Interop.Visio\14.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Dll d:\PIA
```



