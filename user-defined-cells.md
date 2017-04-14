# User-Defined Cells

### Getting the user-defined cells

```
// For a Single shape
var udcell1s = VA.Shapes.UserDefinedCellsHelper.Get(shape1);

// For multiple shapes
var props = VA.Shapes.UserDefinedCellsHelper.Get (page1, shapes);
```

### Adding a new cell

```
VA.Shapes.UserDefinedCellsHelper.Set (s1, prop.Name, prop.Value, prop.Prompt);
```

### Finding out how many user defined cells exist

```
int count = VA.Shapes.UserDefinedCellsHelper.GetCount(s1);
```

### Checking if a cell exists

```
bool exists = VA.Shapes.UserDefinedCellsHelper.Contains(s1, "FOO1");
```

### Deleting a cell

```
VA.Shapes.UserDefinedCellsHelper.Delete(s1, "FOO3");
```



