#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
#endregion

namespace MaterialTransfer
{
  [Transaction(TransactionMode.Manual)]
  public class Command : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Application app = uiapp.Application;
      Document doc = uidoc.Document;

      // get all doc materials
      Dictionary<string, Material> docMaterials = new FilteredElementCollector(doc)
        .OfClass(typeof(Material))
        .Cast<Material>()
        .ToDictionary(m => m.Name);

      var sheetMaterials = GetSheetMaterials();

      //      try
      //      {
      var tx = new Transaction(doc);
      tx.Start("Добавление материалов");
      foreach (var sheetMaterial in sheetMaterials.Values)
      {
        if (sheetMaterial[1] is string sheetMaterialName)
        {
          if (!docMaterials.ContainsKey(sheetMaterialName))
          {
            // Create new material to doc
            var newMaterial = Material.Create(doc, sheetMaterialName);
            var material = doc.GetElement(newMaterial) as Material;

            var matDescription = material.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
            matDescription.Set(sheetMaterial[0] as string);

            var matComments = material.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            if (sheetMaterial[4] is string sheetMaterialComments)
            {
              matComments.Set(sheetMaterialComments);
            }

            // adding color
            material.Color = new Color(243, 23, 54);
            material.Transparency = 0;
            material.UseRenderAppearanceForShading = true;
          }
          else
          {
            Debug.WriteLine($"Material {sheetMaterialName} needs update");
          }
        }
      }
      TaskDialog.Show("Трансфер материалов", $"Обновлено {sheetMaterials.Values.Count} материалов");
      tx.Commit();
      //      }

      //      catch (Exception ex)
      //      {
      //        TaskDialog.Show("Error while loading materials", ex.Message);
      //        return Result.Failed;
      //      }
      return Result.Succeeded;
    }

    public Dictionary<string, List<object>> GetSheetMaterials()
    {

      var db = new DBConnector();
      var materialsFromDB = db.ReadSheetData("Общий!A:E");

      // convert sheet materials to dict
      var sheetMaterials = new Dictionary<string, List<object>>();

      for (int i = 1; i < materialsFromDB.Count; i++)
      {
        // TODO: add null check if material is correct
        var sheetMaterial = materialsFromDB.ElementAt(i);
        if (sheetMaterial.Count() > 4 || !sheetMaterial.Any(x => (string) x == string.Empty))
        {
          string sheetMaterialName = sheetMaterial.ElementAt(1) as string;
          if (!sheetMaterials.ContainsKey(sheetMaterialName))
          {
            sheetMaterials.Add(sheetMaterialName, materialsFromDB[i].ToList());
          }

        }
      }

      return sheetMaterials;
    }
  }
}
