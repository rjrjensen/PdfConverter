namespace PDF_to_XLSX_Converter
{
    class ColumnParsingSettings
    {
        public ColumnParsingSettings(float columnEndingMeasurement, ColumnParsingProfile columnParsingProfile)
        {
            ColumnEndingMeasurement = columnEndingMeasurement;
            ColumnParsingProfile = columnParsingProfile;
        }
        
        public float ColumnEndingMeasurement { get; }
        public ColumnParsingProfile ColumnParsingProfile { get; }
    }
}