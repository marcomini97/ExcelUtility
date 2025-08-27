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
                                toAdd.Descrizione = lineValue.ToString().Trim();
                                break;

                            case 3:
                                if (lineValue != null)
                                    toAdd.D8 = true;
                                break;

                            case 4:
                                if (lineValue != null)
                                    toAdd.D9 = true;
                                break;

                            case 5:
                                if (lineValue != null)
                                    toAdd.D13 = true;
                                break;

                            case 6:
                                if (lineValue != null)
                                    toAdd.D14 = true;
                                break;

                            case 7:
                                if (lineValue != null)
                                    toAdd.D15 = true;
                                break;

                            case 8:
                                if (lineValue != null)
                                    toAdd.R13 = true;
                                break;

                            case 9:
                                if (lineValue != null)
                                    toAdd.R3 = true;
                                break;

                            case 10:
                                if (lineValue != null)
                                    toAdd.R4 = true;
                                break;

                            case 11:
                                if (lineValue != null)
                                    toAdd.R5 = true;
                                break;

                            case 12:
                                if (lineValue != null)
                                    toAdd.R12 = true;
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

            const string connectionString = "Data Source=;Initial Catalog=;User Id=;Password=;Encrypt=True;Trust Server Certificate=true";
            const string query = "INSERT INTO [dbo].[EER]([EerId],[EerCodice],[EerIsPericoloso],[EerDescrizione],[EerD8],[EerD9],[EerD13],[EerD14],[EerD15],[EerR3],[EerR4],[EerR5],[EerR12],[EerR13],[EerIsSos],[EerNota],[EerUteIns],[EerDatIns])VALUES(@Id,@Codice,@Pericoloso,@Descrizione,@D8,@D9,@D13,@D14,@D15,@R3,@R4,@R5,@R12,@R13,@Sospeso,@Nota,@UteIns,@DatIns)";
            
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
                            cmd.Parameters.AddWithValue("@D8", err.D8);
                            cmd.Parameters.AddWithValue("@D9", err.D9);
                            cmd.Parameters.AddWithValue("@D13", err.D13);
                            cmd.Parameters.AddWithValue("@D14", err.D14);
                            cmd.Parameters.AddWithValue("@D15", err.D15);
                            cmd.Parameters.AddWithValue("@R3", err.R3);
                            cmd.Parameters.AddWithValue("@R4", err.R4);
                            cmd.Parameters.AddWithValue("@R5", err.Pericoloso);
                            cmd.Parameters.AddWithValue("@R12", err.Pericoloso);
                            cmd.Parameters.AddWithValue("@R13", err.Pericoloso);
                            cmd.Parameters.AddWithValue("@Sospeso", err.Sospeso);
                            cmd.Parameters.AddWithValue("@Nota", err.Nota);
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
