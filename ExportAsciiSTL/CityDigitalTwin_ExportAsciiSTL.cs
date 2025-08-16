using Eto.Forms;
using Rhino;
using Rhino.Commands;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace AsciiSTLExporter
{
    public class CityDigitalTwin_ExportAsciiSTL : Rhino.Commands.Command
    {
        public override string EnglishName => "CDT_ExportAsciiSTL";

        // ---------- Robustness helpers ----------
        private const double EPS = 1e-8;

        private static bool IsFinite(double v) => !(double.IsNaN(v) || double.IsInfinity(v));
        private static bool IsFinite(Point3d p) => IsFinite(p.X) && IsFinite(p.Y) && IsFinite(p.Z);
        private static bool IsFinite(Vector3d v) => IsFinite(v.X) && IsFinite(v.Y) && IsFinite(v.Z);

        private static bool TriangleIsDegenerate(Point3d a, Point3d b, Point3d c)
        {
            var ab = b - a;
            var ac = c - a;
            var cross = Vector3d.CrossProduct(ab, ac);
            return !IsFinite(cross) || cross.Length <= 1e-12;
        }

        private static Vector3d ComputeFaceNormal(Point3d a, Point3d b, Point3d c)
        {
            var n = Vector3d.CrossProduct(b - a, c - a);
            if (n.Length > 0) n.Unitize();
            return n;
        }

        private static Point3d SnapY(Point3d p)
        {
            if (Math.Abs(p.Y) < EPS) p.Y = 0;
            return p;
        }

        private Point3d InterpolateY0Safe(Point3d a, Point3d b)
        {
            var dy = (b.Y - a.Y);
            if (Math.Abs(dy) < EPS)
                return new Point3d((a.X + b.X) * 0.5, 0, (a.Z + b.Z) * 0.5);
            double t = (0 - a.Y) / dy;
            return new Point3d(
                a.X + t * (b.X - a.X),
                0,
                a.Z + t * (b.Z - a.Z)
            );
        }

        // Appends one triangle to a StringBuilder, computing its own normal.
        // Returns true if triangle was valid and appended; false if skipped.
        private bool AppendTriangle(StringBuilder sb, Point3d a, Point3d b, Point3d c)
        {
            if (!IsFinite(a) || !IsFinite(b) || !IsFinite(c)) return false;
            if (TriangleIsDegenerate(a, b, c)) return false;

            var n = ComputeFaceNormal(a, b, c);
            if (!IsFinite(n) || n.Length <= 1e-12) return false;

            sb.AppendLine($"  facet normal {n.X} {n.Y} {n.Z}");
            sb.AppendLine("    outer loop");
            sb.AppendLine($"      vertex {a.X} {a.Y} {a.Z}");
            sb.AppendLine($"      vertex {b.X} {b.Y} {b.Z}");
            sb.AppendLine($"      vertex {c.X} {c.Y} {c.Z}");
            sb.AppendLine("    endloop");
            sb.AppendLine("  endfacet");
            return true;
        }

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

            try
            {
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

                var meshParams = new MeshingParameters
                {
                    JaggedSeams = false,
                    RefineGrid = true,
                    SimplePlanes = true,
                    MinimumEdgeLength = 0.2,
                    MaximumEdgeLength = 5.0,
                    GridMinCount = 16,
                    GridMaxCount = 256,
                    Tolerance = 0.01,
                    RelativeTolerance = 0.01
                };

                // Preprocess: apply rotation and center shift
                Transform rotation = new Transform(0);
                rotation.M00 = 1; rotation.M01 = 0; rotation.M02 = 0;
                rotation.M10 = 0; rotation.M11 = 0; rotation.M12 = 1;
                rotation.M20 = 0; rotation.M21 = -1; rotation.M22 = 0;
                rotation.M33 = 1;

                var allMeshes = new List<(Mesh mesh, string category)>();
                foreach (var category in categorizedGeometries.Keys)
                {
                    foreach (var geo in categorizedGeometries[category])
                    {
                        var meshes = ConvertToMeshes(geo, meshParams);
                        if (meshes == null) continue;

                        foreach (var m in meshes)
                        {
                            m.Faces.ConvertQuadsToTriangles();
                            m.UnifyNormals();            // consistent winding
                            m.Normals.ComputeNormals();  // vertex normals
                            m.Transform(rotation);

                            // Face normals must be computed AFTER transform
                            m.FaceNormals.ComputeFaceNormals();

                            // Flip normals for terrain-like categories so Y is up
                            if (category == "grounds" || category == "roads" || category == "waters" || category == "grasses")
                            {
                                for (int i = 0; i < m.Faces.Count; i++)
                                {
                                    var face = m.Faces[i];
                                    if (!face.IsTriangle) continue;

                                    var normal = m.FaceNormals[i];
                                    if (normal.Y < 0)
                                    {
                                        // Reverse winding (swap B and C)
                                        m.Faces[i] = new MeshFace(face.A, face.C, face.B);
                                    }
                                }
                                // Recompute normals to ensure consistency
                                m.FaceNormals.ComputeFaceNormals();
                                m.Normals.ComputeNormals();
                            }

                            allMeshes.Add((m, category));
                        }
                    }
                }

                // Compute center of bounding box (world X/Z recentre; Y unchanged)
                var bbox = BoundingBox.Unset;
                foreach (var (mesh, _) in allMeshes)
                    bbox = BoundingBox.Union(bbox, mesh.GetBoundingBox(true));

                var centerShift = Transform.Translation(-bbox.Center.X, 0, -bbox.Center.Z);

                // Apply center shift (translation won't affect normals)
                foreach (var (mesh, _) in allMeshes)
                    mesh.Transform(centerShift);

                // STL Export
                using (var writer = new StreamWriter(path, false, Encoding.ASCII))
                {
                    var categoryCounters = new Dictionary<string, int>();

                    foreach (var (mesh, category) in allMeshes)
                    {
                        if (!categoryCounters.ContainsKey(category))
                            categoryCounters[category] = 1;

                        string baseName;
                        switch (category)
                        {
                            case "roads":
                                baseName = "highway";
                                break;
                            case "other":
                                baseName = "other";
                                break;
                            default:
                                baseName = category.Substring(0, category.Length - 1);
                                break;
                        }

                        string name = $"{baseName}{categoryCounters[category]++}";

                        var sb = new StringBuilder();
                        int facetCount = 0;

                        for (int faceIndex = 0; faceIndex < mesh.Faces.Count; faceIndex++)
                        {
                            var face = mesh.Faces[faceIndex];
                            if (!face.IsTriangle) continue;

                            var A = SnapY(mesh.Vertices[face.A]);
                            var B = SnapY(mesh.Vertices[face.B]);
                            var C = SnapY(mesh.Vertices[face.C]);

                            // Case 1: Entirely below
                            if (A.Y < 0 && B.Y < 0 && C.Y < 0)
                                continue;

                            // Case 2: Entirely above
                            if (A.Y >= 0 && B.Y >= 0 && C.Y >= 0)
                            {
                                if (AppendTriangle(sb, A, B, C)) facetCount++;
                                continue;
                            }

                            // Case 3: Intersecting y=0 plane
                            var above = new List<Point3d>();
                            var below = new List<Point3d>();

                            foreach (var pt in new[] { A, B, C })
                            {
                                if (pt.Y >= 0) above.Add(pt);
                                else below.Add(pt);
                            }

                            if (above.Count == 1 && below.Count == 2)
                            {
                                var p0 = above[0];
                                var i1 = InterpolateY0Safe(p0, below[0]);
                                var i2 = InterpolateY0Safe(p0, below[1]);
                                if (AppendTriangle(sb, p0, i1, i2)) facetCount++;
                            }
                            else if (above.Count == 2 && below.Count == 1)
                            {
                                var p0 = above[0];
                                var p1 = above[1];
                                var i0 = InterpolateY0Safe(p0, below[0]);
                                var i1 = InterpolateY0Safe(p1, below[0]);
                                if (AppendTriangle(sb, p0, p1, i0)) facetCount++;
                                if (AppendTriangle(sb, p1, i1, i0)) facetCount++;
                            }
                        }

                        if (facetCount > 0)
                        {
                            writer.WriteLine($"solid {name}");
                            writer.Write(sb.ToString());
                            writer.WriteLine($"endsolid {name}");
                        }
                        else
                        {
                            RhinoApp.WriteLine($"⏭  Skipped empty solid '{name}' (0 facets after clipping).");
                        }
                    }
                }

                RhinoApp.WriteLine($"Exported ASCII STL to: {path}");
                return Result.Success;
            }
            finally
            {
                // Ensure the temporary document is closed/disposed
                tempDoc?.Dispose();
            }
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
