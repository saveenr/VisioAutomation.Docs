### Dropping Shapes

These samples rely on some extension methods so use this at the top:

```
using VisioAutomation.Extensions;
```

### Dropping one Shape at a time using a master

```
var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
var rectmaster = stencil.Masters["Rectangle"];
var page = doc.Pages.Add();

var shape1 = page.Drop(rectmaster, 1.0, 2.0);

var p = new VisioAutomation.Drawing.Point(5.0, 4.0);
var shape2 = page.Drop(rectmaster, p);
```

### Dropping Multiple Shapes

```
var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
var rectmaster = stencil.Masters["Rectangle"];
var page = doc.Pages.Add();

var centerpoints = new[] {
                               new VisioAutomation.Drawing.Point(1, 2),
                               new VisioAutomation.Drawing.Point(5, 4)
                           };
var masters = new[] { rectmaster, rectmaster };
short[] shapeids = page.DropManyU(masters, centerpoints);
```



