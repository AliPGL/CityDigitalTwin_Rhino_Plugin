using Eto.Forms;
using Rhino;
using Rhino.Commands;
using Rhino.DocObjects;
using Rhino.Geometry;
using Rhino.Input;
using System;
using System.Collections.Generic;
using System.IO;
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

            // Step 1: Identify all ancestor layers named "Buildings"
            var buildingsRootLayers = new List<Layer>();
            foreach (var layer in tempDoc.Layers)
            {
                var current = layer;
                while (current != null)
                {
                    if (current.Name.Trim().Equals("Buildings", StringComparison.OrdinalIgnoreCase))
                    {
                        buildingsRootLayers.Add(layer);
                        break;
                    }
                    current = tempDoc.Layers.FindId(current.ParentLayerId);
                }
            }

            if (buildingsRootLayers.Count == 0)
            {
                RhinoApp.WriteLine("No ancestor layer named 'Buildings' was found.");
                return Result.Nothing;
            }

            // Step 2: Collect all objects under those layers
            var validObjects = new List<GeometryBase>();

            foreach (var obj in tempDoc.Objects)
            {
                var objLayer = tempDoc.Layers[obj.Attributes.LayerIndex];

                foreach (var rootLayer in buildingsRootLayers)
                {
                    // Check if object's layer is a child of one of the Buildings layers
                    var current = objLayer;
                    while (current != null)
                    {
                        if (current.Id == rootLayer.Id)
                        {
                            RhinoApp.WriteLine($"✔ Found building object on layer: {objLayer.FullPath}");
                            CollectGeometryRecursive(obj.Geometry, tempDoc, validObjects, Transform.Identity);
                            break;
                        }
                        current = tempDoc.Layers.FindId(current.ParentLayerId);
                    }
                }
            }

            if (validObjects.Count == 0)
            {
                RhinoApp.WriteLine("No valid building geometry found to export.");
                return Result.Nothing;
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
                int counter = 1;

                foreach (var geo in validObjects)
                {
                    Mesh[] meshes = ConvertToMeshes(geo, meshParams);
                    if (meshes == null || meshes.Length == 0)
                        continue;

                    foreach (var m in meshes)
                    {
                        m.Faces.ConvertQuadsToTriangles();
                        m.Normals.ComputeNormals();

                        writer.WriteLine($"solid building{counter++}");

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

                        writer.WriteLine($"endsolid building{counter - 1}");
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
                            // Accumulate transformation chain
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
            {
                return new[] { mesh.DuplicateMesh() };
            }
            else if (geo is Brep brep)
            {
                return Mesh.CreateFromBrep(brep.DuplicateBrep(), mp);
            }
            else if (geo is Extrusion extrusion)
            {
                var brepGeo = extrusion.ToBrep();
                return Mesh.CreateFromBrep(brepGeo, mp);
            }
            else if (geo is Surface surface)
            {
                var brepGeo = surface.ToBrep();
                return Mesh.CreateFromBrep(brepGeo, mp);
            }

            return null;
        }
    }
}
