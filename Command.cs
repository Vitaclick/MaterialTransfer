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


      var db = new DBConnector();
      var materialsFromDB = db.ReadSheetData("Общий!A:E");
      try
      {
        var tx = new Transaction(doc);
        tx.Start("Добавление материалов");
        for (var i = 1; i < materialsFromDB.Count; i++)
        {
          var dbMaterial = materialsFromDB[i];
          var dbMatName = dbMaterial[1] as string;
          if (dbMatName != null)
          {
            if (!docMaterials.Keys.ToList().Contains(dbMatName))
            {
              // Create new material to doc
              var newMaterial = Material.Create(doc, dbMatName);
              var material = doc.GetElement(newMaterial) as Material;
              var matDescription = material.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION);
              matDescription.Set(dbMaterial[0] as string);
              var matComments = material.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
              try
              {
                matComments.Set(dbMaterial[4] as string);
              }
              catch
              {
                continue;
              }
            }
          }
        }
        tx.Commit();
      }
      
      catch (Exception ex)
      {
        TaskDialog.Show("Error while loading materials", ex.Message);
        return Result.Failed;
      }
      return Result.Succeeded;
    }
  }
}
