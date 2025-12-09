using ExcelUtilityLibrary;
using ExcelUtilityLibrary.Eer;

namespace ExcelUtilityConsole
{
    public class UtilitySettings
    {
        public static void ManageMenuApplication(AziendaSettings aziendaSettings)
        {
            Console.WriteLine("Scegliere il servizio:\n" +
                "1 - Export EER in database" + 
                "\n\n00 per uscire.\n");

            var inputText = Console.ReadLine();

            while (inputText != "00") {

                switch (inputText)
                {
                    case "1":
                        EerExcelService.ReadAndExportEerList(aziendaSettings);
                        break;

                    default:
                        Console.WriteLine("Inserimento non valido");
                        break;
                }

                Console.Write("Effettua la tua scelta: ");
                inputText = Console.ReadLine();
            }
        }
    }
}
