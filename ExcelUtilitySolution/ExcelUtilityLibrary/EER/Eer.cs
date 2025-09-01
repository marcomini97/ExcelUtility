namespace ExcelUtilityLibrary.Eer
{
    public class Eer
    {
        public Guid Id { get; set; }
        public string Codice { get; set; } = string.Empty;
        public bool Pericoloso { get; set; }
        public string Descrizione { get; set; } = string.Empty;
        public int Stato { get; set; }
        public bool AnalisiPreliminare { get; set; }
        public bool Sospeso { get; set; } = false;
        public string Nota { get; set; } = null;
        public string UteIns { get; set; } = string.Empty;
        public DateTime DatIns { get; set; }
    }
}
