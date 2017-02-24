### Custom Properties

### Basic Usage with a single shape

```
var page = doc.Pages.Add();
var s1 = page.DrawRectangle(0, 0, 1, 1);
var cp = new VisioAutomation.Shapes.CustomProperties.CustomPropertyCells();
cp.Value = "Hello World";
VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Set(s1, "Propname", cp);

// Retrieve all the Custom properties from a shape

var props = VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Get(s1);

// Delete the property from the shape

VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Delete(s1, "Propname");
```

NOTES:

* CustomPropertyHelper.Get returns a dictionary. The keys of the dictionary are the custom property names, the values are CustomPropertyCells objects.
* The CustomPropertyCells Object holds the values for the custom properties:
  * Calendar
  * Format
  * Invisible
  * Label
  * LangId
  * Prompt
  * SortKey
  * Type
  * Value
  * Verify

### Working with multiple shapes

```
var page = doc.Pages.Add();
var s1 = page.DrawRectangle(0, 0, 1, 1);
var s2 = page.DrawRectangle(2, 2, 4, 4);

var cp1 = new VisioAutomation.Shapes.CustomProperties.CustomPropertyCells();
cp1.Value = "Hello";
VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Set(s1, "Propname", cp1);

var cp2 = new VisioAutomation.Shapes.CustomProperties.CustomPropertyCells();
cp2.Value = "World";
VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Set(s2, "Propname", cp2);

// Retrieve all the Custom properties from multiple shapes

var shapes = new[] {s1, s2};
var props = VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Get(page,shapes);

// Delete the properties from the shapes
VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Delete(s1, "Propname");
VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Delete(s2, "Propname");
```

### Counting properties and finding names

```
var page = doc.Pages.Add();
var s1 = page.DrawRectangle(0, 0, 1, 1);

var cp1 = new VisioAutomation.Shapes.CustomProperties.CustomPropertyCells();
cp1.Value = "Hello";
VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Set(s1, "Propname", cp1);

int num_custom_props = VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.GetCount(s1);
var custom_prop_names = VisioAutomation.Shapes.CustomProperties.CustomPropertyHelper.Ge
```



