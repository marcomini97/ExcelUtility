using Microsoft.Data.SqlClient;
using OfficeOpenXml;

namespace ExcelUtilityLibrary.Eer
{
    public class EerExcelService
    {
        public static void ReadAndExportEerList(AziendaSettings aziendaSettings)
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), $"File\\{aziendaSettings.NomeFile}");

            Console.WriteLine("Ricerca del documento in corso...");

            if (!File.Exists(filePath))
            {
                Console.WriteLine("\nDocumento non trovato!\n");
                return;
            }

            var listaCodiciEer = _GetEERFromExcel(filePath, aziendaSettings.NomeFoglio);

            _ExportEerList(listaCodiciEer, aziendaSettings.ConnectionString);

            Console.WriteLine("IMPORTAZIONE COMPLETATA\n");
        }

        private static List<Eer> _GetEERFromExcel(string filePath, string nomeFoglio)
        {
            var result = new List<Eer>();

            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets[nomeFoglio];
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
                                if (lineValue != null)
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

                            case 5:
                                if (lineValue != null)
                                    toAdd.Nota1 = true;
                                break;
                            case 6:
                                if (lineValue != null)
                                    toAdd.Nota2 = true;
                                break;
                            case 7:
                                if (lineValue != null)
                                    toAdd.Nota3 = true;
                                break;
                            case 8:
                                if (lineValue != null)
                                    toAdd.Nota4 = true;
                                break;
                            case 9:
                                if (lineValue != null)
                                    toAdd.Nota5 = true;
                                break;

                            case 10:
                                if (lineValue != null)
                                    toAdd.Nota6 = true;
                                break;
                            case 11:
                                if (lineValue != null)
                                    toAdd.Nota7 = true;
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

        private static void _ExportEerList(List<Eer> eerList, string connectionString, bool isDebug = true)
        {
            if (!isDebug)
                return;

            if (string.IsNullOrWhiteSpace(connectionString))
                return;

            const string query = @"INSERT INTO [dbo].[EER]
                                        ([EerId],
                                         [EerCodice],
                                         [EerIsPericoloso],
                                         [EerDescrizione],
                                         [EerStato],
                                         [EerAnalisiPreliminare],
                                         [EerNota1],
                                         [EerNota2],
                                         [EerNota3],
                                         [EerNota4],
                                         [EerNota5],
                                         [EerNota6],
                                         [EerNota7],
                                         [EerIsSos],
                                         [EerUteIns],
                                         [EerDatIns])
                                        VALUES
                                         (@Id,
                                         @Codice,
                                         @Pericoloso,
                                         @Descrizione,
                                         @Stato,
                                         @AnalisiPreliminare,
                                         @Nota1,
                                         @Nota2,
                                         @Nota3,
                                         @Nota4,
                                         @Nota5,
                                         @Nota6,
                                         @Nota7,
                                         @Sospeso,
                                         @UteIns,
                                         @DatIns)";

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
                            cmd.Parameters.AddWithValue("@Nota1", err.Nota1);
                            cmd.Parameters.AddWithValue("@Nota2", err.Nota2);
                            cmd.Parameters.AddWithValue("@Nota3", err.Nota3);
                            cmd.Parameters.AddWithValue("@Nota4", err.Nota4);
                            cmd.Parameters.AddWithValue("@Nota5", err.Nota5);
                            cmd.Parameters.AddWithValue("@Nota6", err.Nota6);
                            cmd.Parameters.AddWithValue("@Nota7", err.Nota7);
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
