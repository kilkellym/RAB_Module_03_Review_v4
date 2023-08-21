#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;

#endregion

namespace RAB_Module_03_Review_v4
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // 1. get furniture data
            List<string[]> furnitureTypeList = GetFurnitureTypes();
            List<string[]> furnitureSetList = GetFurnitureSets();
            
            // 2. remove header row
            furnitureTypeList.RemoveAt(0);
            furnitureSetList.RemoveAt(0);

            // 3. populate furniture data classes
            List<FurnitureType> furnitureTypes = new List<FurnitureType>();

            foreach (string[] curFurnTypeArray in furnitureTypeList)
            {
                FurnitureType curFurnType = new FurnitureType(curFurnTypeArray[0],
                    curFurnTypeArray[1], curFurnTypeArray[2]);

                furnitureTypes.Add(curFurnType);
            }

            // 4. populate furniture set classes
            List<FurnitureSet> furnitureSets = new List<FurnitureSet>();

            foreach (string[] curFurnSetArray in furnitureSetList)
            {
                FurnitureSet curFurnSet = new FurnitureSet(curFurnSetArray[0],
                    curFurnSetArray[1], curFurnSetArray[2]);

                furnitureSets.Add(curFurnSet);
            }

            // 5. get rooms from model
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);

            // 6. loop through rooms
            using (Transaction t = new Transaction(doc))
            {
                int counter = 0;

                t.Start("Move in furniture");

                foreach(SpatialElement curRoom in collector)
                {
                    // 7. get room data
                    LocationPoint roomPoint = curRoom.Location as LocationPoint;
                    XYZ insPoint = roomPoint.Point;

                    string furnSet = GetParameterValueAsString(curRoom, "Furniture Set");

                    // 8. loop through furniture set data - refactor to create GetFurnitureSet method
                    foreach(FurnitureSet curSet in  furnitureSets)
                    {
                        if(curSet.Set == furnSet)
                        {
                            foreach(string furnItem in curSet.Furniture)
                            {
                                foreach(FurnitureType curType in furnitureTypes) // refactor to create GetFurnitureType method
                                {
                                    if(furnItem.Trim() == curType.Name)
                                    {
                                        FamilySymbol curFS = GetFamilySymbolByName(doc, curType.FamilyName, curType.TypeName);
                                    
                                        if(curFS != null)
                                        {
                                            if(curFS.IsActive == false)
                                            {
                                                curFS.Activate();
                                            }
                                        }

                                        FamilyInstance curFI = doc.Create.NewFamilyInstance(insPoint, curFS,
                                            Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                        counter++;
                                    
                                    }
                                }
                            }

                            SetParameterValue(curRoom, "Furniture Count", curSet.GetFurnitureCount());
                            
                        }
                    }
                }

                t.Commit();

                TaskDialog.Show("Complete", $"Inserted {counter} furniture instances");
            }


            return Result.Succeeded;
        }

        private void SetParameterValue(Element curElem, string paramName, int value)
        {
            foreach(Parameter curParam in curElem.Parameters)
            {
                if(curParam.Definition.Name == paramName)
                {
                    curParam.Set(value);
                }
            }
        }
        private void SetParameterValue(Element curElem, string paramName, string value)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                {
                    curParam.Set(value);
                }
            }
        }
        private void SetParameterValue(Element curElem, string paramName, double value)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                {
                    curParam.Set(value);
                }
            }
        }
        private void SetParameterValue(Element curElem, string paramName, ElementId value)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                {
                    curParam.Set(value);
                }
            }
        }

        private FamilySymbol GetFamilySymbolByName(Document doc, string familyName, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));

            foreach(FamilySymbol curFS in collector)
            {
                if(curFS.FamilyName == familyName && curFS.Name == typeName)
                {
                    return curFS;
                }
            }

            return null;
        }

        private string GetParameterValueAsString(Element curElem, string paramName)
        {
            foreach(Parameter curParam in curElem.Parameters)
            {
                if(curParam.Definition.Name == paramName)
                {
                    return curParam.AsString();
                }
            }

            return null;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }

        private List<string[]> GetFurnitureTypes()
        {
            List<string[]> returnList = new List<string[]>();
            returnList.Add(new string[] { "Furniture Name", "Revit Family Name", "Revit Family Type" });
            returnList.Add(new string[] { "desk", "Desk", "60in x 30in" });
            returnList.Add(new string[] { "task chair", "Chair-Task", "Chair-Task" });
            returnList.Add(new string[] { "side chair", "Chair-Breuer", "Chair-Breuer" });
            returnList.Add(new string[] { "bookcase", "Shelving", "96in x 12in x 84in" });
            returnList.Add(new string[] { "loveseat", "Sofa", "54in" });
            returnList.Add(new string[] { "teacher desk", "Table-Rectangular", "48in x 30in" });
            returnList.Add(new string[] { "student desk", "Desk", "60in x 30in Student" });
            returnList.Add(new string[] { "computer desk", "Table-Rectangular", "48in x 30in" });
            returnList.Add(new string[] { "lab desk", "Table-Rectangular", "72in x 30in" });
            returnList.Add(new string[] { "lounge chair", "Chair-Corbu", "Chair-Corbu" });
            returnList.Add(new string[] { "coffee table", "Table-Coffee", "30in x 60in x 18in" });
            returnList.Add(new string[] { "sofa", "Sofa-Corbu", "Sofa-Corbu" });
            returnList.Add(new string[] { "dining table", "Table-Dining", "30in x 84in x 22in" });
            returnList.Add(new string[] { "dining chair", "Chair-Breuer", "Chair-Breuer" });
            returnList.Add(new string[] { "stool", "Chair-Task", "Chair-Task" });

            return returnList;
        }

        private List<string[]> GetFurnitureSets()
        {
            List<string[]> returnList = new List<string[]>();
            returnList.Add(new string[] { "Furniture Set", "Room Type", "Included Furniture" });
            returnList.Add(new string[] { "A", "Office", "desk, task chair, side chair, bookcase" });
            returnList.Add(new string[] { "A2", "Office", "desk, task chair, side chair, bookcase, loveseat" });
            returnList.Add(new string[] { "B", "Classroom - Large", "teacher desk, task chair, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk" });
            returnList.Add(new string[] { "B2", "Classroom - Medium", "teacher desk, task chair, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk" });
            returnList.Add(new string[] { "C", "Computer Lab", "computer desk, computer desk, computer desk, computer desk, computer desk, computer desk, task chair, task chair, task chair, task chair, task chair, task chair" });
            returnList.Add(new string[] { "D", "Lab", "teacher desk, task chair, lab desk, lab desk, lab desk, lab desk, lab desk, lab desk, lab desk, stool, stool, stool, stool, stool, stool, stool" });
            returnList.Add(new string[] { "E", "Student Lounge", "lounge chair, lounge chair, lounge chair, sofa, coffee table, bookcase" });
            returnList.Add(new string[] { "F", "Teacher's Lounge", "lounge chair, lounge chair, sofa, coffee table, dining table, dining chair, dining chair, dining chair, dining chair, bookcase" });
            returnList.Add(new string[] { "G", "Waiting Room", "lounge chair, lounge chair, sofa, coffee table" });

            return returnList;
        }
    }
}
