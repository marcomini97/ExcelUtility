using OfficeOpenXml;

namespace ExcelUtilityLibrary
{
    public class ExcelSettingsLibrary
    {
        public static void SetLicenceKey(string licenceKey)
        {
            ExcelPackage.License.SetNonCommercialPersonal(licenceKey);
        }
    }
}
