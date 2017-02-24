### Identifying Cells

You probably know that Visio has two ways of identifying a cell.

* Cells are identified by names - like "FillForegnd"
* Cells are also identified by a triple of indices - \(a section index, a row index, and a cell index\) - ... or \(s,r,c\)for short

These \(s,r,c\) values are used a lot when you need to work with the visio ShapeSheet. They are so often passed around together that VisioAutomation contains a special struct called SRC to store them. The full name of the struct is VisioAutomation.ShapeSheet.SRC and in this documentation they will always be referred to as the SRC or \(s,r,c\) value

let's create a struct from some literal numbers

```
using VA=VisioAutomation;

var src_0 = new VA.ShapeSheet.SRC(0,0,0);
var src_1 = new VA.ShapeSheet.SRC((short) 0,(short) 0,(short) 0);
```

But, what about a SRC to points to something useful like the fill foreground color? Simple enough ...

```
using SEC=Microsoft.Office.Interop.Visio.VisSectionIndices;
using ROW= Microsoft.Office.Interop.Visio.VisRowIndices;
using CEL= Microsoft.Office.Interop.Visio.VisCellIndices;
using VA=VisioAutomation;

var fg_src = VA.ShapeSheet.SRC(SEC.visSectionObject, ROW.visRowFill, CEL.visFillForegnd);
```

That can be a lot to remember. The good news is VisioAutomation comes with a class that has predefined values for many of these \(s,r,c\) values for cells.

Look at the static class called VisioAutomation.ShapeSheet.SRCConstans there are hundreds of cells identified there. Here are some commonly used ones:

```
Char_Color 
Char_Font 
Char_Size 
FillForegnd 
FillForegndTrans 
FillPattern 
FillBkgnd 
FillBkgndTrans 
LinePattern 
LineWeight 
LineColor 
Height 
Width 
PinX 
PinY 
```

so to get the \(s,r,c\) value for FillPattern we just need to write this

```
using VA=VisioAutomation;

var filpat_src = VA.ShapeSheet.SRCConstants.FillPattern;
```

### Immutability

SRC values are immutable structs. Once created their value cannot be changed

Dealing with the same cell in Multiple rows Often you may have a \(s,r,c\) value , and you need almost the same cell, but for a different row. use the SRC.ForRow method to construct a new \(s,r,c\) value from an existing one var cc\_src\_row0 = VA.ShapeSheet.SRCConstants.Char\_Color; var cc\_src\_row1\_b = cc\_src\_row0.ForRow\(1\);

Converting between cell names and \(s,r,c\) values Suppose you have a cell name stored as a string - "PinX" for example - and now you need to get the \(s,r,c\) value. There are a static methods that will provide the conversion for you for many common cell that are referred to by names. The first is ShapeSheetHelper.GetSRCFromName var pinx\_src = VA.ShapeSheet.ShapeSheetHelper.GetSRCFromName\("PinX"\);

If the cell name is not one the method handles, then an exception will be thrown. So, as an alternative you can try the ShapeSheetHelperTryGetSRCFromName which returns a nullbable SRC value;

```
var pinx_src = VA.ShapeSheet.ShapeSheetHelper.TryGetSRCFromName("PinX");
if (pinx_src.HasValue)
{
    // do nothing
}
```

Converting between \(s,r,c\) values and cell names There's no helper method in the library that will do this. It will be considered for a future version.

