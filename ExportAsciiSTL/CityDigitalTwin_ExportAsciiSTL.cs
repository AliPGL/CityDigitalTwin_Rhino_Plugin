using Eto.Forms;
using Rhino;
using Rhino.Commands;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;

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

        class SolidRecord
        {
            public Mesh Mesh;
            public string Category;   // buildings, trees, grasses, waters, grounds, roads
            public string SolidName;  // the exact STL "solid" name
            public Dictionary<string, string> Props; // effective user text
            public Guid SourceId;
        }

        class GeoItem
        {
            public GeometryBase Geo;
            public Dictionary<string, string> Props;
            public Guid SourceId;
        }

        // Object user text
        static Dictionary<string, string> GetObjectUserText(Rhino.DocObjects.ObjectAttributes a)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (a == null) return d;

            // Read Attribute User Text (the panel you used)
            var keys = a.GetUserStrings();           // enumerates as objects in some RhinoCommon builds
            if (keys != null)
            {
                foreach (var keyObj in keys)
                {
                    if (keyObj is string k)
                    {
                        var val = a.GetUserString(k);
                        if (!string.IsNullOrWhiteSpace(val))
                            d[k] = val;
                    }
                }
            }

            // (Optional) If you also want to honor any ArchivableDictionary entries
            // left by scripts/plugins, merge them with lower precedence:
            var ud = a.UserDictionary;               // ArchivableDictionary
            if (ud != null)
            {
                foreach (var k in ud.Keys)
                {
                    if (!d.ContainsKey(k))
                    {
                        if (ud.TryGetString(k, out string v) && !string.IsNullOrWhiteSpace(v))
                            d[k] = v;
                        else if (ud.ContainsKey(k) && ud[k] != null)
                            d[k] = ud[k].ToString();
                    }
                }
            }

            return d;
        }


        // Layer user text
        static Dictionary<string, string> GetLayerUserText(Rhino.RhinoDoc doc, int layerIndex)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (layerIndex < 0) return d;

            var layer = doc.Layers[layerIndex];
            var ud = layer?.UserDictionary;          // also ArchivableDictionary
            if (ud == null) return d;

            foreach (var key in ud.Keys)
            {
                if (ud.TryGetString(key, out string val) && !string.IsNullOrWhiteSpace(val))
                    d[key] = val;
                else if (ud.ContainsKey(key) && ud[key] != null)
                    d[key] = ud[key].ToString();
            }
            return d;
        }

        // precedence: Object > Layer
        static Dictionary<string, string> EffectiveProps(Rhino.RhinoDoc doc, Rhino.DocObjects.RhinoObject obj)
        {
            var res = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            void take(Dictionary<string, string> src)
            {
                foreach (var kv in src)
                    if (!res.ContainsKey(kv.Key)) res[kv.Key] = kv.Value;
            }
            take(GetObjectUserText(obj.Attributes));
            take(GetLayerUserText(doc, obj.Attributes.LayerIndex));
            return res;
        }

        static string FirstOrDefault(Dictionary<string, string> p, string key)
        {
            return (p.TryGetValue(key, out var v) && !string.IsNullOrWhiteSpace(v)) ? v : "default";
        }

        // Quote a CSV field if it contains commas, quotes, or newlines
        static string Csv(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            bool needsQuotes = s.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0;
            if (!needsQuotes) return s;
            return "\"" + s.Replace("\"", "\"\"") + "\"";
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

                var categorizedGeometries = new Dictionary<string, List<GeoItem>>();
                foreach (var cat in allCategories)
                    categorizedGeometries[cat] = new List<GeoItem>();
                categorizedGeometries["other"] = new List<GeoItem>();

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

                    var props = EffectiveProps(tempDoc, obj);

                    if (props.Count > 0)
                    {
                        RhinoApp.WriteLine($"🔎 Props for {obj.Id}: " +
                            string.Join("; ", props.Select(kv => kv.Key + "=" + kv.Value)));
                    }
                    else
                    {
                        RhinoApp.WriteLine($"⚠️ No user text found for {obj.Id} (layer {tempDoc.Layers[obj.Attributes.LayerIndex].FullPath})");
                    }

                    if (category != null && allCategories.Contains(category))
                    {
                        RhinoApp.WriteLine($"✔ Found {category} object on layer: {tempDoc.Layers[obj.Attributes.LayerIndex].FullPath}");
                        CollectGeometryRecursive(obj.Geometry, tempDoc, categorizedGeometries[category], Transform.Identity, props, obj.Id);
                    }
                    else
                    {
                        RhinoApp.WriteLine($"⚠️ Found unclassified object on layer: {tempDoc.Layers[obj.Attributes.LayerIndex].FullPath} → assigned to 'other'");
                        CollectGeometryRecursive(obj.Geometry, tempDoc, categorizedGeometries["other"], Transform.Identity, props, obj.Id);
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

                var grouped = new Dictionary<(string Category, Guid SourceId), SolidRecord>();

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

                foreach (var category in categorizedGeometries.Keys)
                {
                    foreach (var item in categorizedGeometries[category])
                    {
                        var meshes = ConvertToMeshes(item.Geo, meshParams);
                        if (meshes == null) continue;

                        // one record per (category, source object)
                        var key = (category, item.SourceId);
                        if (!grouped.TryGetValue(key, out var rec))
                        {
                            rec = new SolidRecord
                            {
                                Category = category,
                                Mesh = new Mesh(),      // start empty, we will Append()
                                Props = item.Props,
                                SourceId = item.SourceId
                            };
                            grouped[key] = rec;
                        }

                        // Prepare each partial mesh, then append into the merged one
                        foreach (var m in meshes)
                        {
                            m.Faces.ConvertQuadsToTriangles();
                            m.Transform(rotation);
                            rec.Mesh.Append(m);
                        }
                    }
                }

                // At this point, grouped.Values are the merged solids
                var solidRecords = grouped.Values.ToList();

                // Compute center of bounding box (world X/Z recentre; Y unchanged)
                var bbox = BoundingBox.Unset;
                foreach (var rec in solidRecords)
                    bbox = BoundingBox.Union(bbox, rec.Mesh.GetBoundingBox(true));

                var centerShift = Transform.Translation(-bbox.Center.X, 0, -bbox.Center.Z);

                // Apply center shift (translation won't affect normals)
                foreach (var rec in solidRecords)
                    rec.Mesh.Transform(centerShift);

                // Recompute normals after merging + shifting
                foreach (var rec in solidRecords)
                {
                    var m = rec.Mesh;

                    // Optional polish
                    m.Compact();                   // clean unused vertices
                    m.Weld(0.0174533);              // ~1 degree weld to smooth seams

                    m.UnifyNormals();
                    m.Normals.ComputeNormals();
                    m.FaceNormals.ComputeFaceNormals();

                    if (rec.Category == "grounds" || rec.Category == "roads" ||
                        rec.Category == "waters" || rec.Category == "grasses")
                    {
                        for (int i = 0; i < m.Faces.Count; i++)
                        {
                            var face = m.Faces[i];
                            if (!face.IsTriangle) continue;
                            var normal = m.FaceNormals[i];
                            if (normal.Y < 0)
                                m.Faces[i] = new MeshFace(face.A, face.C, face.B); // flip winding
                        }
                        m.FaceNormals.ComputeFaceNormals();
                        m.Normals.ComputeNormals();
                    }
                }

                var csvVeg = new List<string> { "solidName,vegetationType,soilType" };
                var csvRoad = new List<string> { "solidName,roadType" };
                var csvWater = new List<string> { "solidName,waterType" };
                var csvBldg = new List<string> { "buildingID,isUsageTypeKnown,usageType,isHeightClassKnown,heightClassification,isConstructionYearKnown,constructionYear,scheduleType,envelopePropertyType,thermalZoneModel,hasHVAC,isHVACSystemTypeKnown,heatingSystemType,coolingSystemType,isHotwaterSystemTypeKnown,hotwaterSystemType" };

                // STL Export
                using (var writer = new StreamWriter(path, false, Encoding.ASCII))
                {
                    var categoryCounters = new Dictionary<string, int>();

                    var baseNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["buildings"] = "building",
                        ["trees"] = "tree",
                        ["grasses"] = "grass",
                        ["waters"] = "waterway",
                        ["grounds"] = "ground",
                        ["roads"] = "highway",
                        ["other"] = "building", // convert all "other" to building
                    };

                    foreach (var rec in solidRecords)
                    {
                        var category = rec.Category;
                        if (!categoryCounters.ContainsKey(category))
                            categoryCounters[category] = 1;

                        var baseName = baseNameMap.TryGetValue(category, out var mapped) ? mapped : "building";
                        rec.SolidName = $"{baseName}{categoryCounters[category]++}";

                        var mesh = rec.Mesh;
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
                            writer.WriteLine($"solid {rec.SolidName}");
                            writer.Write(sb.ToString());
                            writer.WriteLine($"endsolid {rec.SolidName}");

                            var p = rec.Props;

                            RhinoApp.WriteLine($"CSV Export for {rec.SolidName} ({rec.Category}) → " + string.Join("; ", p.Select(kv => kv.Key + "=" + kv.Value)));

                            // vegetation (trees/grasses/grounds) → solidVegetationSoilSubcategories.csv
                            if (rec.Category == "trees")
                                csvVeg.Add($"{Csv(rec.SolidName)},{Csv(FirstOrDefault(p, "CDT.tree_type"))},{Csv(FirstOrDefault(p, "CDT.soil_type"))}");
                            else if (rec.Category == "grasses")
                                csvVeg.Add($"{Csv(rec.SolidName)},{Csv(FirstOrDefault(p, "CDT.grass_type"))},{Csv(FirstOrDefault(p, "CDT.soil_type"))}");
                            else if (rec.Category == "grounds")
                                csvVeg.Add($"{Csv(rec.SolidName)},{Csv(FirstOrDefault(p, "CDT.ground_type"))},{Csv(FirstOrDefault(p, "CDT.soil_type"))}");

                            // roads → solidRoadSubcategories.csv
                            else if (rec.Category == "roads")
                                csvRoad.Add($"{Csv(rec.SolidName)},{Csv(FirstOrDefault(p, "CDT.road_type"))}");

                            // waters → solidWaterSubcategories.csv
                            else if (rec.Category == "waters")
                                csvWater.Add($"{Csv(rec.SolidName)},{Csv(FirstOrDefault(p, "CDT.water_type"))}");

                            // buildings → building_info_basic.csv
                            else if (rec.Category == "buildings")
                            {
                                var usage = FirstOrDefault(p, "CDT.building_type");   // usageType
                                var year = FirstOrDefault(p, "CDT.building_year");   // constructionYear
                                var isUsage = usage != "default" ? "yes" : "no";
                                var isYear = year != "default" ? "yes" : "no";

                                csvBldg.Add(string.Join(",",
                                    Csv(rec.SolidName),  // buildingID
                                    Csv(isUsage), Csv(usage),
                                    Csv("no"), Csv("default"),
                                    Csv(isYear), Csv(year),
                                    Csv("default"), Csv("default"),
                                    Csv("default"),
                                    Csv("default"), Csv("default"),
                                    Csv("default"), Csv("default"),
                                    Csv("default"), Csv("default")
                                ));
                            }
                        }
                        else
                        {
                            RhinoApp.WriteLine($"⏭  Skipped empty solid '{rec.SolidName}' (0 facets after clipping).");
                        }
                    }
                }

                RhinoApp.WriteLine($"Exported ASCII STL to: {path}");

                var outDir = Path.GetDirectoryName(path);
                var utf8Bom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);

                File.WriteAllLines(Path.Combine(outDir, "solidVegetationSoilSubcategories.csv"), csvVeg, utf8Bom);
                File.WriteAllLines(Path.Combine(outDir, "solidRoadSubcategories.csv"), csvRoad, utf8Bom);
                File.WriteAllLines(Path.Combine(outDir, "solidWaterSubcategories.csv"), csvWater, utf8Bom);
                File.WriteAllLines(Path.Combine(outDir, "building_info_basic.csv"), csvBldg, utf8Bom);

                return Result.Success;
            }
            finally
            {
                // Ensure the temporary document is closed/disposed
                tempDoc?.Dispose();
            }
        }

        private void CollectGeometryRecursive(GeometryBase geo, RhinoDoc doc, List<GeoItem> result, Transform accumulatedTransform, Dictionary<string, string> props, Guid sourceId)
        {
            if (geo is InstanceReferenceGeometry instanceRef)
            {
                var instanceDef = doc.InstanceDefinitions.Find(instanceRef.ParentIdefId, true);
                var instanceXform = instanceRef.Xform;
                if (instanceDef != null)
                {
                    foreach (var o in instanceDef.GetObjects())
                    {
                        if (o.Geometry != null)
                        {
                            var xform = accumulatedTransform * instanceXform;
                            CollectGeometryRecursive(o.Geometry, doc, result, xform, props, sourceId);
                        }
                    }
                }
            }
            else if (geo is Mesh || geo is Brep || geo is Extrusion || geo is Surface)
            {
                var dup = geo.Duplicate();
                dup.Transform(accumulatedTransform);
                result.Add(new GeoItem { Geo = dup, Props = props, SourceId = sourceId });
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
