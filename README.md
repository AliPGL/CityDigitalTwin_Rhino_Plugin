# Rhino ASCII STL Exporter

This Rhino plugin allows users to automatically export 3D geometry from a `.3dm` file into a single **ASCII STL** file. It supports classification by layer name and organizes exported geometry into categories such as `buildings`, `trees`, `roads`, etc.

## 🚀 Features

- ✅ Automatically loads a Rhino `.3dm` file
- ✅ Recursively extracts geometry from nested block instances
- ✅ Detects and classifies geometry based on ancestor layer names:
  - `Buildings`
  - `Trees`
  - `Grasses`
  - `Waters`
  - `Grounds`
  - `Roads`
- ✅ Generates STL solids with names like:
  - `building1`, `building2`, ...
  - `tree1`, `tree2`, ...
  - `other1`, `other2`, ... (for unclassified geometry)
- ✅ High-quality mesh conversion with custom meshing parameters

## 📂 Layer Naming Convention

To ensure proper classification, place your geometry under layers named with one of the following keywords:

- `Buildings`
- `Trees`
- `Grasses`
- `Waters`
- `Grounds`
- `Roads`

The plugin will recursively check all **ancestor layers** of each object to determine its category.

If no matching category is found, the object is added under `"other"`.

## 🔧 Usage

1. Open Rhino.
2. Run the command: `CDT_ExportAsciiSTL`
3. Select the `.3dm` file to export.
4. Choose the output `.stl` file path.
5. The plugin will:
   - Load the model
   - Traverse objects
   - Classify geometry
   - Generate ASCII STL with solid names

## 📦 Output

- One ASCII `.stl` file containing all valid geometry.
- Each `solid` block in the STL will be named according to its category.

Example:

```stl
solid building1
  facet normal ...
  ...
endsolid building1

solid tree1
  ...
endsolid tree1
🛠 Development Notes
Built with C# using the RhinoCommon SDK

Compatible with Rhino 7 and Rhino 8

Works with nested blocks and complex layer hierarchies

📬 Feedback & Contributions
Have a feature request or want to contribute? Feel free to reach out or fork this project!
```
