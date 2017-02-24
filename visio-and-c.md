## Introduction

It's useful to understand some how some tasks are done in Visio directly using the automation - without using the VisioAutomation library.



## Hello World in C\# \(VS2015\)

### Add a reference to the Visio Primary Interop Assembly:

* In the Solution Explorer, right click on References, select Add Reference
* The Add Reference dialog will launch
* In the .NET Tab select Microsoft.Office.Interop.Visio
* Click OK

Sample code:

```
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using IVisio = Microsoft.Office.Interop.Visio;
namespace Visio2007AutomationHelloWorldCSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            var visapp = new IVisio.Application();
            var doc = visapp.Documents.Add("");
            var page = visapp.ActivePage;
            var shape = page.DrawRectangle(1, 1, 5, 4);
            shape.Text = "Hello World";
        }
    }
}


// Add a new empty doc
var doc = app.Documents.Add( "" );



```

## High-Performance ShapeSheet operations

### 



