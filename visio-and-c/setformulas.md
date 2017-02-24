### Page.SetFormulas

```
public static void Page_SetFormulas(IVisio.Document doc)
{
    var pages = doc.Pages;
    var page = pages.Add();
    page.NameU = "PSF";

    var shape = page.DrawRectangle(1, 1, 4, 3);
    shape.get_CellsU("Width").Formula = "=(1.0+2.5)";
    shape.get_CellsU("Height").Formula = "=(0.0+1.5)";

    // BUILD UP THE REQUEST
    short flags = (short)(IVisio.VisGetSetArgs.visSetBlastGuards 
                        | IVisio.VisGetSetArgs.visSetUniversalSyntax);
    var items = new[]
    {
        new {   shapeid = (short) shape.ID,
                section = (short) IVisio.VisSectionIndices.visSectionObject, 
                row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                cell    = (short) IVisio.VisCellIndices.visXFormWidth,
                formula = "1.3" },        

        new {   shapeid = (short) shape.ID,
                section = (short) IVisio.VisSectionIndices.visSectionObject, 
                row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                cell    = (short) IVisio.VisCellIndices.visXFormHeight,
                formula = "7.71" }
    };

    // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
    short[] SID_SRCStream = new short[items.Length * 4];
    object[] formulas_objects = new object[items.Length];
    for (int i = 0; i < items.Length; i++)
    {
        SID_SRCStream[i * 4 + 0] = items[i].shapeid;
        SID_SRCStream[i * 4 + 1] = items[i].section;
        SID_SRCStream[i * 4 + 2] = items[i].row;
        SID_SRCStream[i * 4 + 3] = items[i].cell;
        formulas_objects[i] = items[i].formula;
    }

    // EXECUTE THE REQUEST
    short flags = (short)(IVisio.VisGetSetArgs.visSetBlastGuards | IVisio.VisGetSetArgs.visSetUniversalSyntax);
    int count = page.SetFormulas(SID_SRCStream, formulas_objects, flags);

    // DISPLAY THE INFORMATION
    shape.Text = "SetFormulas";
}
```

### Shape.SetFormulas

```
public static void Shape_SetFormulas(IVisio.Document doc)
{
    var pages = doc.Pages;
    var page = pages.Add();
    page.NameU = "SSF";

    var shape = page.DrawRectangle(1, 1, 4, 3);
    shape.get_CellsU("Width").Formula = "=(1.0+2.5)";
    shape.get_CellsU("Height").Formula = "=(0.0+1.5)";

    // BUILD UP THE REQUEST
    short flags = (short)(IVisio.VisGetSetArgs.visSetBlastGuards 
                        | IVisio.VisGetSetArgs.visSetUniversalSyntax);
    var items = new[]
    {
        new {   section = (short) IVisio.VisSectionIndices.visSectionObject, 
                row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                cell    = (short) IVisio.VisCellIndices.visXFormWidth,
                formula = "1.3" },  

        new {   section = (short) IVisio.VisSectionIndices.visSectionObject, 
                row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                cell    = (short) IVisio.VisCellIndices.visXFormHeight,
                formula = "7.71" }
    };

    // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
    short[] SRCStream = new short[items.Length * 3];
    object[] formulas_objects = new object[items.Length];
    for (int i = 0; i 
<
 items.Length; i++)
    {
        SRCStream[i * 3 + 0] = items[i].section;
        SRCStream[i * 3 + 1] = items[i].row;
        SRCStream[i * 3 + 2] = items[i].cell;
        formulas_objects[i] = items[i].formula;
    }

    // EXECUTE THE REQUEST
    short flags = (short)(IVisio.VisGetSetArgs.visSetBlastGuards | IVisio.VisGetSetArgs.visSetUniversalSyntax);
    int count = shape.SetFormulas(SRCStream, formulas_objects, flags);

    // DISPLAY THE INFORMATION
    shape.Text = "SetFormulas";
}
```



