# User-Defined Cells

### Getting the user-defined cells

```
// For a Single shape
var udcell1s = VA.Shapes.UserDefinedCells.UserDefinedCellsHelper.Get(shape1);

// For multiple shapes
var props = VA.Shapes.UserDefinedCells.UserDefinedCellsHelper.Get (page1, shapes);

```

### Adding a new cell

```
VA.UserDefinedCells.UserDefinedCellsHelper.Set (s1, prop.Name, prop.Value, prop.Prompt);

```

### Finding out how many user defined cells exist

```
int count = VA.UserDefinedCells.UserDefinedCellsHelper.GetCount(s1);

```

### Checking if a cell exists

```
bool exists = VA.UserDefinedCells.UserDefinedCellsHelper.Contains(s1, "FOO1");

```

### Deleting a cell

```
VA.UserDefinedCells.UserDefinedCellsHelper.Delete(s1, "FOO3");
```



