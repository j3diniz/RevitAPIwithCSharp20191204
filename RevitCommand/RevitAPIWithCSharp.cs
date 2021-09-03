using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace RevitCommand {
    [Transaction(TransactionMode.Manual)]
    class RevitAPIWithCSharp : IExternalCommand {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {

            // Establishing connections
            UIApplication application = commandData.Application;
            UIDocument uiDocument = application.ActiveUIDocument;
            Document document = uiDocument.Document;
            Selection selection = uiDocument.Selection;

            using (Transaction transactionExtrusion = new Transaction(document, "Creating a new Extrusion")) {
                transactionExtrusion.Start();
                try {

                    SketchPlane sketchPlane;
                    sketchPlane = CreateSketchPlaneFromUIDocument(uiDocument);
                    CreateExtrusion(document, sketchPlane);

                    transactionExtrusion.Commit();

                    return Result.Succeeded;

                } catch (Exception ex) {
                    TaskDialog.Show("Revit API Sample fail!", ex.Message);
                    transactionExtrusion.RollBack();
                    return Result.Failed;
                }
            }
        }

        /// <summary>
        /// Create Sketch Plane from Application
        /// </summary>
        /// <param name="application"></param>
        /// <returns></returns>
        private SketchPlane CreateSketchPlaneFromApplication(UIApplication application) {
            //try to create a new sketch plane
            XYZ newNormal = new XYZ(1, 1, 0);  // The normal vector
            XYZ newOrigin = new XYZ(0, 0, 0);  // The origin point

            Plane geometryPlane = Plane.CreateByNormalAndOrigin(newNormal, newOrigin); // Create geometry plane

            // Create sketch plane
            SketchPlane sketchPlane = SketchPlane.Create(application.ActiveUIDocument.Document, geometryPlane);

            return sketchPlane;
        }

        /// <summary>
        /// Create Sketch Plane from uiDocument
        /// </summary>
        /// <param name="uiDocument"></param>
        /// <returns></returns>
        private SketchPlane CreateSketchPlaneFromUIDocument(UIDocument uiDocument) {
            Reference faceReference = uiDocument.Selection.PickObject(ObjectType.Face);

            GeometryObject geoObject = uiDocument.Document.GetElement(faceReference).GetGeometryObjectFromReference(faceReference);
            PlanarFace planarFace = geoObject as PlanarFace;

            XYZ newNormal = planarFace.FaceNormal;
            XYZ newOrigin = planarFace.Origin;

            Plane geometryPlane = Plane.CreateByNormalAndOrigin(newNormal, newOrigin); // Create geometry plane

            // Create sketch plane
            SketchPlane sketchPlane = SketchPlane.Create(uiDocument.Document, geometryPlane);

            return sketchPlane;
        }


        /// <summary>
        /// Create Extrusion
        /// </summary>
        /// <param name="document"></param>
        /// <param name="sketchPlane"></param>
        /// <returns></returns>
        private Extrusion CreateExtrusion(Autodesk.Revit.DB.Document document, SketchPlane sketchPlane) {
            Extrusion rectExtrusion = null;

            // make sure we have a family document
            if (true == document.IsFamilyDocument) {
                // define the profile for the extrusion
                CurveArrArray curveArrArray = new CurveArrArray();
                CurveArray curveArray1 = new CurveArray();
                CurveArray curveArray2 = new CurveArray();
                CurveArray curveArray3 = new CurveArray();

                // create a rectangular profile
                XYZ p0 = XYZ.Zero;
                XYZ p1 = new XYZ(5, 0, 0);
                XYZ p2 = new XYZ(5, 5, 0);
                XYZ p3 = new XYZ(0, 5, 0);
                Line line1 = Line.CreateBound(p0, p1);
                Line line2 = Line.CreateBound(p1, p2);
                Line line3 = Line.CreateBound(p2, p3);
                Line line4 = Line.CreateBound(p3, p0);
                curveArray1.Append(line1);
                curveArray1.Append(line2);
                curveArray1.Append(line3);
                curveArray1.Append(line4);

                curveArrArray.Append(curveArray1);

                // create solid rectangular extrusion
                rectExtrusion = document.FamilyCreate.NewExtrusion(true, curveArrArray, sketchPlane, 5);

                if (null != rectExtrusion) {
                    // move extrusion to proper place
                    XYZ transPoint1 = new XYZ(-16, 0, 0);
                    ElementTransformUtils.MoveElement(document, rectExtrusion.Id, transPoint1);
                } else {
                    throw new Exception("Create new Extrusion failed.");
                }
            } else {
                throw new Exception("Please open a Family document before invoking this command.");
            }

            return rectExtrusion;
        }

    }

}
