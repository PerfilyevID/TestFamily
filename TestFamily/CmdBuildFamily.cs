using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TestFamily.Common;
using View = Autodesk.Revit.DB.View;

namespace TestFamily
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CmdBuildFamily : IExternalCommand
    {
        private static readonly double WIDTH = 500.ToFeets();
        private static readonly string[] TEMPLATES = new string[]
        {
            @"C:\ProgramData\Autodesk\RVT {0}\Family Templates\Russian\Метрическая система, типовая модель.rft",
            @"C:\ProgramData\Autodesk\RVT {0}\Family Templates\English\Metric Generic Model.rft",
            @"C:\ProgramData\Autodesk\RVT {0}\Family Templates\German\Allgemeines Modell.rft",
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string source = string.Format(TEMPLATES.First(x => File.Exists(string.Format(x, commandData.Application.Application.VersionNumber))), commandData.Application.Application.VersionNumber);
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Revit families (*.rfa)|*.rfa";
            saveDialog.RestoreDirectory = true;
            saveDialog.CheckFileExists = false;
            saveDialog.CheckPathExists = true;
            saveDialog.Title = "Select file to create family";
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                string fullPath = saveDialog.FileName;
                Document doc = commandData.Application.Application.NewFamilyDocument(source);
                doc.SaveAs(ModelPathUtils.ConvertUserVisiblePathToModelPath(fullPath), new SaveAsOptions() { OverwriteExistingFile = true, MaximumBackups = 1 });
                doc = commandData.Application.OpenAndActivateDocument(fullPath).Document;
                using (Transaction t = new Transaction(doc, "Create family"))
                {
                    t.Start();

                    #region Geometry
                    XYZ p0 = new XYZ(-WIDTH / 2, -WIDTH / 2, 0);
                    XYZ p1 = new XYZ(WIDTH / 2, -WIDTH / 2, 0);
                    XYZ p2 = new XYZ(WIDTH / 2, WIDTH / 2, 0);
                    XYZ p3 = new XYZ(-WIDTH / 2, WIDTH / 2, 0);

                    CurveArray curveArray = new CurveArray();
                    curveArray.Append(Line.CreateBound(p0, p1));
                    curveArray.Append(Line.CreateBound(p1, p2));
                    curveArray.Append(Line.CreateBound(p2, p3));
                    curveArray.Append(Line.CreateBound(p3, p0));

                    CurveArrArray curveArrArray = new CurveArrArray();
                    curveArrArray.Append(curveArray);
                    #endregion

                    Level defaultLevel = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements().First() as Level;
                    SketchPlane levelPlane = new FilteredElementCollector(doc).OfClass(typeof(SketchPlane))
                        .Cast<SketchPlane>()
                        .Where(s => s.IsSuitableForModelElements)
                        .First(e => e.Name.Equals(defaultLevel.Name));

                    Extrusion extrusion = doc.FamilyCreate.NewExtrusion(true, curveArrArray, levelPlane, WIDTH);
                    FamilyParameter familyParameterWidth = doc.FamilyManager.AddParameter("w", BuiltInParameterGroup.PG_GEOMETRY, ParameterType.Length, true);
                    View defaultPlan = doc.GetElement(defaultLevel.FindAssociatedPlanViewId()) as View;
                    List<Line> lines = new List<Line>();
                    foreach (Line l in extrusion.Sketch.Profile.get_Item(0)) lines.Add(l);
                    CreateDimension(doc, lines[1], lines[3], defaultPlan, centerReference: GetCentralReferencePlaneByNormal(doc, XYZ.BasisY));
                    CreateDimension(doc, lines[1], lines[3], defaultPlan, familyParameterWidth);
                    CreateDimension(doc, lines[0], lines[2], defaultPlan, centerReference: GetCentralReferencePlaneByNormal(doc, XYZ.BasisX));
                    CreateDimension(doc, lines[0], lines[2], defaultPlan, familyParameterWidth);
                    doc.FamilyManager.AssociateElementParameterToFamilyParameter(extrusion.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM), familyParameterWidth);
                    t.Commit();
                }
                doc.Save(new SaveOptions() { Compact = false });
                return Result.Succeeded;
            }
            return Result.Cancelled;
        }

        private static void CreateDimension(Document doc, Curve a, Curve b, View view, FamilyParameter param=null, Reference centerReference=null)
        {
            bool isMultiple = centerReference != null;
            ReferenceArray referencaArray = new ReferenceArray();
            referencaArray.Append(a.Reference);
            if(isMultiple) referencaArray.Append(centerReference);
            referencaArray.Append(b.Reference);
            Line line = Line.CreateBound(a.GetEndPoint(0), b.GetEndPoint(1));
            XYZ direction = new XYZ(-line.Direction.Y, line.Direction.X, 0);
            line = Line.CreateBound(line.GetEndPoint(0) + direction * (isMultiple ? 2 : 1), line.GetEndPoint(1) + direction * (isMultiple ? 2 : 1));
            Dimension dimension = doc.FamilyCreate.NewLinearDimension(view, line, referencaArray);
            if(!isMultiple && param != null)
                dimension.FamilyLabel = param;
            else dimension.AreSegmentsEqual = true;
        }

        private static Reference GetCentralReferencePlaneByNormal(Document doc, XYZ normal)
        {
            foreach (ReferencePlane referencePlane in new FilteredElementCollector(doc).OfClass(typeof(ReferencePlane)))
            {
                Plane plane = referencePlane.GetPlane();
                if ((plane.Normal.IsAlmostEqualTo(normal) || plane.Normal.IsAlmostEqualTo(normal.Negate())) && plane.Origin.IsAlmostEqualTo(XYZ.Zero))
                    return referencePlane.GetReference();
            }
            throw new NullReferenceException("Reference plane not found");
        }
    }
}
