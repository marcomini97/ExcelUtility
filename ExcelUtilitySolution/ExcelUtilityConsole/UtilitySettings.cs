using ExcelUtilityLibrary.Eer;

namespace ExcelUtilityConsole
{
    public class UtilitySettings
    {
        public static void ManageMenuApplication()
        {
            Console.WriteLine("Scegliere il servizio:\n" +
                "1 - Export EER in database" + 
                "\n\n00 per uscire.\n");

            var stringProva = Console.ReadLine();

            while (stringProva != "00") {

                switch (stringProva)
                {
                    case "1":
                        EerExcelService.ReadAndExportEerList();
                        break;
                }

                stringProva = Console.ReadLine();
            }
        }
    }
}
