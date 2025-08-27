namespace ExcelUtilityLibrary.Eer
{
    public class Eer
    {
        public Guid Id { get; set; }
        public string Codice { get; set; } = string.Empty;
        public bool Pericoloso { get; set; }
        public string Descrizione { get; set; } = string.Empty;
        public bool D8 { get; set; } = false;
        public bool D9 { get; set; } = false;
        public bool D13 { get; set; } = false;
        public bool D14 { get; set; } = false;
        public bool D15 { get; set; } = false;
        public bool R3 { get; set; } = false;
        public bool R4 { get; set; } = false;
        public bool R5 { get; set; } = false;
        public bool R12 { get; set; } = false;
        public bool R13 { get; set; } = false;
        public bool Sospeso { get; set; } = false;
        public string Nota { get; set; } = string.Empty;
        public string UteIns { get; set; } = string.Empty;
        public DateTime DatIns { get; set; }
    }
}
