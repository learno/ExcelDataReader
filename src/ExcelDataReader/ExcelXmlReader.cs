namespace ExcelDataReader;

internal sealed class ExcelXmlReader : ExcelDataReader<Core.XmlFormat.XmlWorkbook, Core.XmlFormat.XmlWorksheet>
{
    public ExcelXmlReader(System.IO.Stream stream)
    {
        Workbook = new Core.XmlFormat.XmlWorkbook(stream);

        // By default, the data reader is positioned on the first result.
        Reset();
    }

    public override void Close()
    {
        base.Close();
        Workbook = null;
    }
}

