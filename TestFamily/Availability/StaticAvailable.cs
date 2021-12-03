using Autodesk.Revit.UI;

namespace TestFamily.Availability
{
    public class StaticAvailable : IExternalCommandAvailability
    {
        public bool IsCommandAvailable(UIApplication applicationData, Autodesk.Revit.DB.CategorySet selectedCategories) => true;
    }
}
