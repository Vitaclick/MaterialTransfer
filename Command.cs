#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

      var db = new DBConnector();
      var materialsFromDB = db.ReadSheetData("Общий!A:C");

      try
      {
        foreach (var dbMaterial in materialsFromDB)
        {
          var dbMatName = dbMaterial[0] as string;
          // Create new material to doc
          var newMaterial = Material.Create(doc, dbMatName);
          var material = doc.GetElement(newMaterial) as Material;

        }
      }
      
      catch (Exception ex)
      {
        return Result.Failed;
      }
      return Result.Succeeded;
    }
  }
}
