using Microsoft.Data.SqlClient;
using OfficeOpenXml;

namespace ExcelUtilityLibrary.Eer
{
    public class EerExcelService
    {
        public static void ReadAndExportEerList()
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "File\\CodiciEER.xlsx");

            Console.WriteLine("Ricerca del documento in corso...");

            if (!File.Exists(filePath))
            {
                Console.WriteLine("\nDocumento non trovato!\n");
                return;
            }

            var listaCodiciEer = _GetEERFromExcel(filePath);

            _ExportEerList(listaCodiciEer);

            Console.WriteLine("IMPORTAZIONE COMPLETATA\n");
        }
        
        private static List<Eer> _GetEERFromExcel(string filePath)
        {
            var result = new List<Eer>();

            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets["Foglio1"];
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                for (int row = 3; row <= rowCount; row++)
                {
                    var toAdd = new Eer();

                    for (int col = 1; col <= colCount; col++)
                    {
                        var lineValue = worksheet.Cells[row, col].Value;

                        switch (col)
                        {
                            case 1:
                                var stringa = lineValue.ToString();

                                if (stringa.Contains('*'))
                                {
                                    stringa = stringa.Remove(8);
                                    toAdd.Pericoloso = true;
                                }
                                toAdd.Codice = stringa.Replace(" ", "");
                                break;

                            case 2:
                                stringa = lineValue.ToString();

                                if (string.IsNullOrWhiteSpace(lineValue.ToString()))
                                    toAdd.Descrizione = "VALORE NON PRESENTE";

                                else
                                    toAdd.Descrizione = char.ToUpper(stringa[0]) + stringa.Substring(1);
                                break;

                            case 3:
                                if(lineValue != null)
                                {
                                    if (lineValue == "A")
                                        toAdd.Stato = 2;
                                    else
                                        toAdd.Stato = 1;
                                }

                                break;

                            case 4:
                                if (lineValue != null)
                                    toAdd.AnalisiPreliminare = true;
                                break;
                        }

                    }

                    toAdd.Id = Guid.NewGuid();
                    toAdd.UteIns = "SERVER";
                    toAdd.DatIns = DateTime.Now;

                    result.Add(toAdd);
                }
            }

            return result;
        }

        private static void _ExportEerList(List<Eer> eerList, bool isDebug = true)
        {
            if (!isDebug)
                return;

            const string connectionString = "Data Source=NB-MCOMINI;Initial Catalog=DbProduzioneRMB;User Id=sa;Password=Abcd.1234;Encrypt=True;Trust Server Certificate=true";
            const string query = "INSERT INTO [dbo].[EER]([EerId],[EerCodice],[EerIsPericoloso],[EerDescrizione],[EerStato],[EerAnalisiPreliminare],[EerIsSos],[EerUteIns],[EerDatIns])VALUES(@Id,@Codice,@Pericoloso,@Descrizione,@Stato,@AnalisiPreliminare,@Sospeso,@UteIns,@DatIns)";
            
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                try
                {
                    conn.Open();

                    foreach (var err in eerList)
                    {
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@Id", err.Id);
                            cmd.Parameters.AddWithValue("@Codice", err.Codice);
                            cmd.Parameters.AddWithValue("@Pericoloso", err.Pericoloso);
                            cmd.Parameters.AddWithValue("@Descrizione", err.Descrizione);
                            cmd.Parameters.AddWithValue("@Stato", err.Stato);
                            cmd.Parameters.AddWithValue("@AnalisiPreliminare", err.AnalisiPreliminare);
                            cmd.Parameters.AddWithValue("@Sospeso", err.Sospeso);
                            cmd.Parameters.AddWithValue("@UteIns", err.UteIns);
                            cmd.Parameters.AddWithValue("@DatIns", err.DatIns);

                            int rowsAffected = cmd.ExecuteNonQuery();
                        }
                    }

                    conn.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Errore: " + ex.Message);
                }
            }
        }
    }
}
