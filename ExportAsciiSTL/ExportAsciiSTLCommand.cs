using Eto.Forms;
using Rhino;
using Rhino.Commands;
using Rhino.DocObjects;
using Rhino.Geometry;
using Rhino.Input;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace AsciiSTLExporter
{
    public class ExportAsciiSTLCommand : Rhino.Commands.Command
    {
        public override string EnglishName => "ExportAsciiSTL";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            var openDialog = new OpenFileDialog
            {
                Title = "Select a 3DM file to export to ASCII STL",
                Filters = { new FileFilter("Rhino 3D Files", ".3dm") }
            };

            if (openDialog.ShowDialog(null) != DialogResult.Ok)
                return Result.Cancel;

            string filePath = openDialog.FileName;
            var tempDoc = RhinoDoc.Open(filePath, out var errors);
            if (tempDoc == null)
            {
                RhinoApp.WriteLine("Failed to open the 3DM file.");
                return Result.Failure;
            }

            RhinoApp.WriteLine($"Loaded model from: {filePath}");

            var categories = new[] { "buildings", "trees", "grasses", "waters", "grounds", "roads" };
            var allCategories = new HashSet<string>(categories);
            var layerCategoryMap = new Dictionary<Guid, string>();

            foreach (var layer in tempDoc.Layers)
            {
                var current = layer;
                while (current != null)
                {
                    var lname = current.Name.Trim().ToLower();
                    if (allCategories.Contains(lname))
                    {
                        layerCategoryMap[layer.Id] = lname;
                        break;
                    }
                    current = tempDoc.Layers.FindId(current.ParentLayerId);
                }
            }

            var categorizedGeometries = new Dictionary<string, List<GeometryBase>>();
            foreach (var cat in allCategories)
                categorizedGeometries[cat] = new List<GeometryBase>();
            categorizedGeometries["other"] = new List<GeometryBase>();

            foreach (var obj in tempDoc.Objects)
            {
                var layer = tempDoc.Layers[obj.Attributes.LayerIndex];
                string category = null;

                while (layer != null)
                {
                    if (layerCategoryMap.TryGetValue(layer.Id, out category))
                        break;

                    layer = tempDoc.Layers.FindId(layer.ParentLayerId);
                }

                if (category != null && allCategories.Contains(category))
                {
                    RhinoApp.WriteLine($"✔ Found {category} object on layer: {tempDoc.Layers[obj.Attributes.LayerIndex].FullPath}");
                    CollectGeometryRecursive(obj.Geometry, tempDoc, categorizedGeometries[category], Transform.Identity);
                }
                else
                {
                    RhinoApp.WriteLine($"⚠️ Found unclassified object on layer: {tempDoc.Layers[obj.Attributes.LayerIndex].FullPath} → assigned to 'other'");
                    CollectGeometryRecursive(obj.Geometry, tempDoc, categorizedGeometries["other"], Transform.Identity);
                }
            }

            var saveDialog = new Rhino.UI.SaveFileDialog
            {
                Filter = "ASCII STL (*.stl)|*.stl",
                DefaultExt = "stl",
                Title = "Save ASCII STL"
            };

            if (!saveDialog.ShowSaveDialog())
                return Result.Cancel;

            var path = saveDialog.FileName;
            if (string.IsNullOrWhiteSpace(path))
                return Result.Cancel;

            var meshParams = MeshingParameters.Default;
            meshParams.JaggedSeams = false;
            meshParams.RefineGrid = true;
            meshParams.SimplePlanes = false;
            meshParams.MinimumEdgeLength = 0.01;
            meshParams.MaximumEdgeLength = 2.0;
            meshParams.GridMinCount = 100;
            meshParams.GridMaxCount = 10000;

            using (var writer = new StreamWriter(path, false, Encoding.ASCII))
            {
                foreach (var category in categorizedGeometries.Keys)
                {
                    var geos = categorizedGeometries[category];
                    int counter = 1;

                    foreach (var geo in geos)
                    {
                        var meshes = ConvertToMeshes(geo, meshParams);
                        if (meshes == null) continue;

                        foreach (var m in meshes)
                        {
                            m.Faces.ConvertQuadsToTriangles();
                            m.Normals.ComputeNormals();

                            string baseName = category == "other" ? "other" : category.Substring(0, category.Length - 1);
                            string name = $"{baseName}{counter++}";
                            writer.WriteLine($"solid {name}");

                            foreach (var face in m.Faces)
                            {
                                if (!face.IsTriangle) continue;

                                var A = m.Vertices[face.A];
                                var B = m.Vertices[face.B];
                                var C = m.Vertices[face.C];

                                var normal = Vector3d.CrossProduct(B - A, C - A);
                                normal.Unitize();

                                writer.WriteLine($"  facet normal {normal.X} {normal.Y} {normal.Z}");
                                writer.WriteLine("    outer loop");
                                writer.WriteLine($"      vertex {A.X} {A.Y} {A.Z}");
                                writer.WriteLine($"      vertex {B.X} {B.Y} {B.Z}");
                                writer.WriteLine($"      vertex {C.X} {C.Y} {C.Z}");
                                writer.WriteLine("    endloop");
                                writer.WriteLine("  endfacet");
                            }

                            writer.WriteLine($"endsolid {name}");
                        }
                    }
                }
            }

            RhinoApp.WriteLine($"Exported ASCII STL to: {path}");
            return Result.Success;
        }

        private void CollectGeometryRecursive(GeometryBase geo, RhinoDoc doc, List<GeometryBase> result, Transform accumulatedTransform)
        {
            if (geo is InstanceReferenceGeometry instanceRef)
            {
                var instanceDef = doc.InstanceDefinitions.Find(instanceRef.ParentIdefId, true);
                var instanceXform = instanceRef.Xform;

                if (instanceDef != null)
                {
                    foreach (var obj in instanceDef.GetObjects())
                    {
                        if (obj.Geometry != null)
                        {
                            var xform = accumulatedTransform * instanceXform;
                            CollectGeometryRecursive(obj.Geometry, doc, result, xform);
                        }
                    }
                }
            }
            else if (geo is Mesh || geo is Brep || geo is Extrusion || geo is Surface)
            {
                var dup = geo.Duplicate();
                dup.Transform(accumulatedTransform);
                result.Add(dup);
            }
        }

        private Mesh[] ConvertToMeshes(GeometryBase geo, MeshingParameters mp)
        {
            if (geo is Mesh mesh)
                return new[] { mesh.DuplicateMesh() };
            else if (geo is Brep brep)
                return Mesh.CreateFromBrep(brep.DuplicateBrep(), mp);
            else if (geo is Extrusion extrusion)
                return Mesh.CreateFromBrep(extrusion.ToBrep(), mp);
            else if (geo is Surface surface)
                return Mesh.CreateFromBrep(surface.ToBrep(), mp);

            return null;
        }
    }
}
