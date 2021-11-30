using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Visual;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.VisualBasic;


namespace RevitMaterialTexture
{
    [Transaction(TransactionMode.Manual)]
    public class MaterialTextureCmd : IExternalApplication
    {
        static UIControlledApplication _cachedUiCtrApp;

        public Result OnStartup(UIControlledApplication application)
        {
            _cachedUiCtrApp = application;

            try
            {
                _cachedUiCtrApp.Idling += OnIdling;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                return Result.Failed;
            }
        }

        private void OnIdling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {
            _cachedUiCtrApp.Idling -= OnIdling;

            UIApplication app = sender as UIApplication;

            string name = "Material Texture Test";
            app.CreateRibbonTab(name);
            RibbonPanel panel = app.CreateRibbonPanel(name, name);

            PushButtonData buttonData = new PushButtonData(name, "Create Materials with Texture", Assembly.GetExecutingAssembly().Location, "RevitMaterialTexture.MaterialTexture");
            PushButton button = null;

            button = panel.AddItem(buttonData) as PushButton;
            button.ToolTip = "Create Materials with Texture";
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    public class MaterialTexture : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string input = Interaction.InputBox("Enter the number of materials/faces to create", "Material Texture Test", "100");

            int totalFaces;

            int.TryParse(input, out totalFaces);

            Document document = commandData.Application.ActiveUIDocument.Document;

            string folderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string templateImagePath = Path.Combine(folderPath, "TestMaterial.jpg");

            Dictionary<ElementId, Material> faces = new Dictionary<ElementId, Material>();

            Transaction transaction = new Transaction(document, "delete old faces");
            transaction.Start();

            // Delete the any previous faces, materials and appearance asset elements

            string namesToCheck = $"Test_Face_";

            var collector = new FilteredElementCollector(document).WhereElementIsNotElementType();

            List<ElementId> idsToDelete = new List<ElementId>();

            var genericModelIds = collector.OfCategory(BuiltInCategory.OST_GenericModel).ToElements().
                Where(e => e.Name.IndexOf(namesToCheck, StringComparison.OrdinalIgnoreCase) >= 0).Select(e => e.Id).ToList();
            var materialIds = collector.OfCategory(BuiltInCategory.OST_Materials).
                Where(e => e.Name.IndexOf(namesToCheck, StringComparison.OrdinalIgnoreCase) >= 0).Select(e => e.Id).ToList();
            var assetIds = collector.OfClass(typeof(AppearanceAssetElement)).
                Where(e => e.Name.IndexOf(namesToCheck, StringComparison.OrdinalIgnoreCase) >= 0).Select(e => e.Id).ToList();

            if (genericModelIds != null) idsToDelete.AddRange(genericModelIds);
            if (materialIds != null) idsToDelete.AddRange(materialIds);
            if (assetIds != null) idsToDelete.AddRange(assetIds);

            document.Delete(idsToDelete);

            document.Regenerate();
            transaction.Commit();

            commandData.Application.ActiveUIDocument.RefreshActiveView();

            transaction = new Transaction(document, "create faces and materials");
            transaction.Start();

            System.Windows.Forms.Application.DoEvents();

            for (int i = 0; i < totalFaces; i++)
            {
                string materialName = $"Test_Face_{i}";

                File.Copy(templateImagePath, Path.Combine(folderPath, materialName + ".jpg"), true);

                // check if the material already exists.
                Material material = new FilteredElementCollector(document).OfClass((typeof(Material))).Cast<Material>().FirstOrDefault(m => m.Name == materialName);

                if (material == null)
                {
                    ElementId materialId = Material.Create(document, materialName);
                    material = document.GetElement(materialId) as Material;

                    material.Color = new Color(125, 125, 125);
                    material.Shininess = 0;
                    material.Transparency = 0;
                    material.UseRenderAppearanceForShading = true;
                }

                // create a new face 3 feet above the previous one
                List<XYZ> localVertices = new List<XYZ>(3);
                localVertices.Add(new XYZ(100, 100, i * 3));
                localVertices.Add(new XYZ(100, 110, i * 3));
                localVertices.Add(new XYZ(110, 110, i * 3));
                localVertices.Add(new XYZ(110, 100, i * 3));

                TessellatedShapeBuilder builder = new TessellatedShapeBuilder();

                builder.OpenConnectedFaceSet(true);

                builder.AddFace(new TessellatedFace(localVertices, material.Id));
                builder.CloseConnectedFaceSet();
                builder.Build();

                TessellatedShapeBuilderResult result = builder.GetBuildResult();

                DirectShape shape = DirectShape.CreateElement(document, new ElementId(BuiltInCategory.OST_GenericModel));
                shape.SetShape(result.GetGeometricalObjects());
                shape.Name = materialName;

                faces.Add(shape.Id, material);
            }

            // commit is needed before material properties are applied
            // so that the faces can be queried as DirectShape
            TransactionStatus status = transaction.Commit();

            transaction = new Transaction(document, "save material");
            transaction.Start();

            System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch();
            watch.Start();

            foreach (var kvp in faces)
            {
                Material material = kvp.Value;

                string materialName = material.Name;

                Dictionary<string, object> textureProperties = new Dictionary<string, object>();
                textureProperties.Add("generic_transparency", 5);
                textureProperties.Add("UnifiedBitmap.TextureRealWorldScaleX", 10d);
                textureProperties.Add("UnifiedBitmap.TextureRealWorldScaleY", 10d);
                textureProperties.Add("UnifiedBitmap.UnifiedbitmapBitmap", Path.Combine(folderPath, materialName + ".jpg"));
                textureProperties.Add("UnifiedBitmap.TextureURepeat", false);
                textureProperties.Add("UnifiedBitmap.TextureVRepeat", false);

                textureProperties.Add("UnifiedBitmap.TextureWAngle", 0d);
                textureProperties.Add("UnifiedBitmap.TextureRealWorldOffsetX", 0d);
                textureProperties.Add("UnifiedBitmap.TextureRealWorldOffsetY", 0d);

                SetTextureProperties(material, document, materialName, textureProperties);
            }

            status = transaction.Commit();
            transaction.Dispose();

            watch.Stop();

            System.Windows.Forms.MessageBox.Show("Texture Material properties applied in " + watch.ElapsedMilliseconds / 1000.0 + " seconds");

            return Result.Succeeded;
        }

        private static AppearanceAssetElement genericAsset;

        private void SetTextureProperties(object material, object document, string matname, Dictionary<string, object> properties)
        {
            try
            {
                Material mat = material as Material;
                Document doc = document as Document;

                if (mat.AppearanceAssetId == ElementId.InvalidElementId)
                {
                    // Not checking if the element exists as they are deleted before hand
                    AppearanceAssetElement assetElement = null; // = (AppearanceAssetElement)(new FilteredElementCollector(doc).OfClass(typeof(AppearanceAssetElement)).Where(a => a.Name == matname).FirstOrDefault());

                    if (assetElement == null) // Avoid copying the template asset per material
                    {
                        try
                        {
                            if (genericAsset == null)
                            {
                                genericAsset = new FilteredElementCollector(doc).OfClass(typeof(AppearanceAssetElement))
                                .ToElements()
                                .Cast<AppearanceAssetElement>().FirstOrDefault(i => i.Name.Contains("Generic"));
                            }

                            if (genericAsset != null)
                            {
                                assetElement = genericAsset.Duplicate(matname);
                            }
                        }
                        catch
                        {
                            // it may be faster to let it crash in case the asset with the same name already exists
                            // than to check for it every time. as they are deleted in advance anyway

                            assetElement = (AppearanceAssetElement)(new FilteredElementCollector(doc).OfClass(typeof(AppearanceAssetElement)).Where(a => a.Name == matname).FirstOrDefault());
                        }
                    }

                    if (assetElement == null) throw new Exception("Could not find Default AppearanceAssetElement to duplicate its Asset");

                    mat.AppearanceAssetId = assetElement.Id;
                }

                using (AppearanceAssetEditScope editScope = new AppearanceAssetEditScope(doc))
                {
                    Asset editableAsset = editScope.Start(mat.AppearanceAssetId);

                    if (editableAsset != null)
                    {
                        if (properties.ElementAt(0).Key == "generic_transparency")
                        {
                            AssetPropertyDouble transparencyProperty = editableAsset.FindByName("generic_transparency") as AssetPropertyDouble;

                            double transparency = Convert.ToDouble(properties.ElementAt(0).Value);

                            if (transparency < 0) transparency = 0;
                            else if (transparency > 100) transparency = 100;

                            transparency /= 100; // value is needed between 0 and 1

                            if (transparencyProperty.IsValidValue(transparency)) // in case generic_transparency property is not present.
                            {
                                transparencyProperty.Value = transparency;
                            }
                        }

                        AssetProperty assetProperty = editableAsset.FindByName("generic_diffuse");

                        if (assetProperty == null) // in case generic diffuse property is not present. // image may be defined under Parameters
                        {
                            int size = editableAsset.Size;
                            for (int assetIdx = 0; assetIdx < size; assetIdx++)
                            {
                                AssetProperty ap = editableAsset.Get(assetIdx);

                                if (ap.NumberOfConnectedProperties < 1) continue;

                                Asset ca = ap.GetConnectedProperty(0) as Asset;
                                if (ca.Name == "UnifiedBitmapSchema")
                                {
                                    assetProperty = ap; break;
                                }
                            }
                        }

                        if (assetProperty.NumberOfConnectedProperties == 0)
                            assetProperty.AddConnectedAsset("UnifiedBitmapSchema");

                        Asset connectedAsset = assetProperty.GetConnectedProperty(0) as Asset;
                        if (connectedAsset.Name == "UnifiedBitmapSchema")
                        {
                            foreach (var kvp in properties)
                            {
                                string key = GetUnifiedBitmapKey(kvp.Key);

                                if (key == "") continue;

                                if (key == UnifiedBitmap.TextureRealWorldScaleX || key == UnifiedBitmap.TextureRealWorldScaleY ||
                                    key == UnifiedBitmap.TextureRealWorldOffsetX || key == UnifiedBitmap.TextureRealWorldOffsetY)
                                {
                                    AssetPropertyDistance assetPD = connectedAsset.FindByName(key) as AssetPropertyDistance;

                                    assetPD.Value = (double)kvp.Value;
                                }
                                else if (key == UnifiedBitmap.TextureOffsetLock)
                                {
                                    AssetPropertyBoolean offLock = connectedAsset.FindByName(key) as AssetPropertyBoolean;
                                    offLock.Value = (bool)kvp.Value;
                                }
                                else if (key == UnifiedBitmap.TextureScaleLock)
                                {
                                    AssetPropertyBoolean scaleLock = connectedAsset.FindByName(key) as AssetPropertyBoolean;
                                    scaleLock.Value = (bool)kvp.Value;
                                }
                                else if (key == UnifiedBitmap.TextureURepeat)
                                {
                                    AssetPropertyBoolean repeatU = connectedAsset.FindByName(key) as AssetPropertyBoolean;
                                    repeatU.Value = (bool)kvp.Value;
                                }
                                else if (key == UnifiedBitmap.TextureVRepeat)
                                {
                                    AssetPropertyBoolean repeatV = connectedAsset.FindByName(key) as AssetPropertyBoolean;
                                    repeatV.Value = (bool)kvp.Value;
                                }
                                else if (key == UnifiedBitmap.TextureWAngle)
                                {
                                    AssetPropertyDouble ang = connectedAsset.FindByName(key) as AssetPropertyDouble;
                                    if (ang.IsValidValue((double)kvp.Value)) ang.Value = (double)kvp.Value;
                                }
                                else if (key == UnifiedBitmap.UnifiedbitmapBitmap)
                                {
                                    AssetPropertyString path = connectedAsset.FindByName(key) as AssetPropertyString;
                                    if (path.IsValidValue(kvp.Value.ToString())) path.Value = kvp.Value.ToString();
                                }
                            }
                        }
                    }

                    editScope.Commit(true);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string GetUnifiedBitmapKey(string key)
        {
            if (key == "UnifiedBitmap.TextureRealWorldScaleX") return UnifiedBitmap.TextureRealWorldScaleX;
            else if (key == "UnifiedBitmap.TextureRealWorldScaleY") return UnifiedBitmap.TextureRealWorldScaleY;
            else if (key == "UnifiedBitmap.TextureRealWorldOffsetX") return UnifiedBitmap.TextureRealWorldOffsetX;
            else if (key == "UnifiedBitmap.TextureRealWorldOffsetY") return UnifiedBitmap.TextureRealWorldOffsetY;
            else if (key == "UnifiedBitmap.UnifiedbitmapBitmap") return UnifiedBitmap.UnifiedbitmapBitmap;
            else if (key == "UnifiedBitmap.TextureURepeat") return UnifiedBitmap.TextureURepeat;
            else if (key == "UnifiedBitmap.TextureVRepeat") return UnifiedBitmap.TextureVRepeat;
            else if (key == "UnifiedBitmap.TextureOffsetLock") return UnifiedBitmap.TextureOffsetLock;
            else if (key == "UnifiedBitmap.TextureScaleLock") return UnifiedBitmap.TextureScaleLock;
            else if (key == "UnifiedBitmap.TextureWAngle") return UnifiedBitmap.TextureWAngle;
            else return key;
        }

    }
}
