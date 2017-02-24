### Page.Page.SetResults

```
 public static void Page_SetResults(IVisio.Document doc)
    {
        var pages = doc.Pages;
        var page = pages.Add();
        page.NameU = "PSR";

        var shape = page.DrawRectangle(1, 1, 4, 3);
        shape.get_CellsU("Width").Formula = "=(1.0+2.5)";
        shape.get_CellsU("Height").Formula = "=(0.0+1.5)";

        // BUILD UP THE REQUEST
        short flags = 0;
        var items = new[]
        {
            new {   shapeid = (short) shape.ID,
                    section = (short) IVisio.VisSectionIndices.visSectionObject, 
                    row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                    cell    = (short) IVisio.VisCellIndices.visXFormWidth,
                    result = 8.0 ,
                    unitcode = (short) IVisio.VisUnitCodes.visNoCast },    

            new {   shapeid = (short) shape.ID,
                    section = (short) IVisio.VisSectionIndices.visSectionObject, 
                    row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                    cell    = (short) IVisio.VisCellIndices.visXFormHeight,
                    result = 1.0 ,
                    unitcode = (short) IVisio.VisUnitCodes.visNoCast }
        };

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        short[] SID_SRCStream = new short[items.Length * 4];
        object[] results_objects = new object[items.Length];
        object[] unitcodes = new object[items.Length];
        for (int i = 0; i 
<
 items.Length; i++)
        {
            SID_SRCStream[i * 4 + 0] = items[i].shapeid;
            SID_SRCStream[i * 4 + 1] = items[i].section;
            SID_SRCStream[i * 4 + 2] = items[i].row;
            SID_SRCStream[i * 4 + 3] = items[i].cell;
            results_objects[i] = items[i].result;
            unitcodes[i] = items[i].unitcode;
        }

        // EXECUTE THE REQUEST
        short flags = (short)(IVisio.VisGetSetArgs.visSetBlastGuards | IVisio.VisGetSetArgs.visSetUniversalSyntax);
        int count = shape.SetFormulas(SRCStream, formulas_objects, flags);

        // DISPLAY THE INFORMATION
        shape.Text = "SetFormulas";

    }

```

### Shape.SetResults

```
public static void Shape_SetResults(IVisio.Document doc)
{
    var pages = doc.Pages;
    var page = pages.Add();
    page.NameU = "SSR";

    var shape = page.DrawRectangle(1, 1, 4, 3);
    shape.get_CellsU("Width").Formula = "=(1.0+2.5)";
    shape.get_CellsU("Height").Formula = "=(0.0+1.5)";

    // BUILD UP THE REQUEST
    short flags = 0;
    var items = new[]
    {
        new {   section = (short) IVisio.VisSectionIndices.visSectionObject, 
                row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                cell    = (short) IVisio.VisCellIndices.visXFormWidth,
                result = 8.0 ,
                unitcode = (short) IVisio.VisUnitCodes.visNoCast },                                

        new {   section = (short) IVisio.VisSectionIndices.visSectionObject,
                row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                cell    = (short) IVisio.VisCellIndices.visXFormHeight,
                result = 1.0 ,
                unitcode = (short) IVisio.VisUnitCodes.visNoCast }
    };

    // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
    short[] SRCStream = new short[items.Length * 3];
    object[] results_objects = new object[items.Length];
    object[] unitcodes = new object[items.Length];
    for (int i = 0; i 
<
 items.Length; i++)
    {
        SRCStream[i * 3 + 0] = items[i].section;
        SRCStream[i * 3 + 1] = items[i].row;
        SRCStream[i * 3 + 2] = items[i].cell;
        results_objects[i] = items[i].result;
        unitcodes[i] = items[i].unitcode;
    }

    // EXECUTE THE REQUEST
    short flags = 0;
    int count = shape.SetResults(SRCStream, unitcodes, results_objects, flags);

    // DISPLAY THE INFORMATION
    shape.Text = "SetResults";
}
}
```



