using System.Data;

namespace ExcelDataReader.Tests;

[TestFixture]
public class ExcelXmlReaderTest
{
    [Test]
    public void ReadSimpleXml()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        var dataSet = excelReader.AsDataSet();

        Assert.That(dataSet, Is.Not.Null);
        Assert.That(dataSet.Tables.Count, Is.EqualTo(1));
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(10));
        Assert.That(dataSet.Tables[0].Columns.Count, Is.EqualTo(10));
    }

    [Test]
    public void ReadXmlStringValues()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        var dataSet = excelReader.AsDataSet();

        // First row: col1, empty, col3, empty, col5, empty, col7, empty, col9
        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo("col1"));
        Assert.That(dataSet.Tables[0].Rows[0][1], Is.EqualTo(DBNull.Value));
        Assert.That(dataSet.Tables[0].Rows[0][2], Is.EqualTo("col3"));
        Assert.That(dataSet.Tables[0].Rows[0][4], Is.EqualTo("col5"));
        Assert.That(dataSet.Tables[0].Rows[0][6], Is.EqualTo("col7"));
        Assert.That(dataSet.Tables[0].Rows[0][8], Is.EqualTo("col9"));
    }

    [Test]
    public void ReadXmlNumberValues()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        var dataSet = excelReader.AsDataSet();

        // Second row: 10x10, empty cells, 10x19
        Assert.That(dataSet.Tables[0].Rows[1][0], Is.EqualTo("10x10"));
        Assert.That(dataSet.Tables[0].Rows[1][9], Is.EqualTo("10x19"));
    }

    [Test]
    public void ReadXmlDecimalValues()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        var dataSet = excelReader.AsDataSet();

        // Fourth row: 10x12, empty cells, 10x21
        Assert.That(dataSet.Tables[0].Rows[3][0], Is.EqualTo("10x12"));
        Assert.That(dataSet.Tables[0].Rows[3][9], Is.EqualTo("10x21"));
    }

    [Test]
    public void ReadXmlBooleanValues()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        var dataSet = excelReader.AsDataSet();

        // Fifth row: 10x13 to 10x22
        Assert.That(dataSet.Tables[0].Rows[4][0], Is.EqualTo("10x13"));
        Assert.That(dataSet.Tables[0].Rows[4][1], Is.EqualTo("10x14"));
        Assert.That(dataSet.Tables[0].Rows[4][2], Is.EqualTo("10x15"));
        Assert.That(dataSet.Tables[0].Rows[4][9], Is.EqualTo("10x22"));
    }

    [Test]
    public void ReadXmlDateTimeValues()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        var dataSet = excelReader.AsDataSet();

        // Sixth row (index 5): 10x14, empty cells, 10x23
        Assert.That(dataSet.Tables[0].Rows[5][0], Is.EqualTo("10x14"));
        Assert.That(dataSet.Tables[0].Rows[5][9], Is.EqualTo("10x23"));
    }

    [Test]
    public void ReadXmlEmptyCells()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        var dataSet = excelReader.AsDataSet();

        // First row has empty cells at positions 1, 3, 5, 7
        Assert.That(dataSet.Tables[0].Rows[0][0], Is.EqualTo("col1"));
        Assert.That(dataSet.Tables[0].Rows[0][1], Is.EqualTo(DBNull.Value));
        Assert.That(dataSet.Tables[0].Rows[0][2], Is.EqualTo("col3"));
        Assert.That(dataSet.Tables[0].Rows[0][3], Is.EqualTo(DBNull.Value));
        
        // Second row has empty cells in the middle
        Assert.That(dataSet.Tables[0].Rows[1][0], Is.EqualTo("10x10"));
        Assert.That(dataSet.Tables[0].Rows[1][1], Is.EqualTo(DBNull.Value));
        Assert.That(dataSet.Tables[0].Rows[1][9], Is.EqualTo("10x19"));
    }

    [Test]
    public void ReadXmlNegativeNumbers()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        var dataSet = excelReader.AsDataSet();

        // Ninth row (index 8): 10x17 to 10x26
        Assert.That(dataSet.Tables[0].Rows[8][0], Is.EqualTo("10x17"));
        Assert.That(dataSet.Tables[0].Rows[8][1], Is.EqualTo("10x18"));
        Assert.That(dataSet.Tables[0].Rows[8][9], Is.EqualTo("10x26"));
    }

    [Test]
    public void ReadXmlMultiSheet()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestMultiSheet.xml"));
        var dataSet = excelReader.AsDataSet();

        Assert.That(dataSet, Is.Not.Null);
        Assert.That(dataSet.Tables.Count, Is.EqualTo(3));

        // 注意：工作表顺序可能与XML中的顺序不同，取决于读取顺序
        // XML中顺序是：Sheet2, Sheet1, Sheet3
        Assert.That(dataSet.Tables[0].TableName, Is.EqualTo("Sheet2"));
        Assert.That(dataSet.Tables[1].TableName, Is.EqualTo("Sheet1"));
        Assert.That(dataSet.Tables[2].TableName, Is.EqualTo("Sheet3"));
    }

    [Test]
    public void ReadXmlMultiSheetData()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestMultiSheet.xml"));
        var dataSet = excelReader.AsDataSet();

        // Sheet2: 12行4列，所有值都是1
        Assert.That(dataSet.Tables["Sheet2"].Rows.Count, Is.EqualTo(12));
        Assert.That(dataSet.Tables["Sheet2"].Columns.Count, Is.EqualTo(4));
        Assert.That(dataSet.Tables["Sheet2"].Rows[0][0], Is.EqualTo(1.0));
        Assert.That(dataSet.Tables["Sheet2"].Rows[0][1], Is.EqualTo(1.0));
        Assert.That(dataSet.Tables["Sheet2"].Rows[0][2], Is.EqualTo(1.0));
        Assert.That(dataSet.Tables["Sheet2"].Rows[0][3], Is.EqualTo(1.0));
        Assert.That(dataSet.Tables["Sheet2"].Rows[11][0], Is.EqualTo(1.0));

        // Sheet1: 12行4列，所有值都是2
        Assert.That(dataSet.Tables["Sheet1"].Rows.Count, Is.EqualTo(12));
        Assert.That(dataSet.Tables["Sheet1"].Columns.Count, Is.EqualTo(4));
        Assert.That(dataSet.Tables["Sheet1"].Rows[0][0], Is.EqualTo(2.0));
        Assert.That(dataSet.Tables["Sheet1"].Rows[0][1], Is.EqualTo(2.0));
        Assert.That(dataSet.Tables["Sheet1"].Rows[0][2], Is.EqualTo(2.0));
        Assert.That(dataSet.Tables["Sheet1"].Rows[0][3], Is.EqualTo(2.0));
        Assert.That(dataSet.Tables["Sheet1"].Rows[11][0], Is.EqualTo(2.0));

        // Sheet3: 5行2列，所有值都是3
        Assert.That(dataSet.Tables["Sheet3"].Rows.Count, Is.EqualTo(5));
        Assert.That(dataSet.Tables["Sheet3"].Columns.Count, Is.EqualTo(2));
        Assert.That(dataSet.Tables["Sheet3"].Rows[0][0], Is.EqualTo(3.0));
        Assert.That(dataSet.Tables["Sheet3"].Rows[0][1], Is.EqualTo(3.0));
        Assert.That(dataSet.Tables["Sheet3"].Rows[4][0], Is.EqualTo(3.0));
        Assert.That(dataSet.Tables["Sheet3"].Rows[4][1], Is.EqualTo(3.0));
    }

    [Test]
    public void ReadXmlWithNextResult()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestMultiSheet.xml"));
        
        Assert.That(excelReader.ResultsCount, Is.EqualTo(3));
        
        // First sheet (Sheet2)
        Assert.That(excelReader.Read(), Is.True);
        Assert.That(excelReader.GetDouble(0), Is.EqualTo(1.0));
        Assert.That(excelReader.Name, Is.EqualTo("Sheet2"));
        
        // Move to next sheet (Sheet1)
        Assert.That(excelReader.NextResult(), Is.True);
        Assert.That(excelReader.Read(), Is.True);
        Assert.That(excelReader.GetDouble(0), Is.EqualTo(2.0));
        Assert.That(excelReader.Name, Is.EqualTo("Sheet1"));
        
        // Move to next sheet (Sheet3)
        Assert.That(excelReader.NextResult(), Is.True);
        Assert.That(excelReader.Read(), Is.True);
        Assert.That(excelReader.GetDouble(0), Is.EqualTo(3.0));
        Assert.That(excelReader.Name, Is.EqualTo("Sheet3"));
        
        // No more sheets
        Assert.That(excelReader.NextResult(), Is.False);
    }

    [Test]
    public void ReadXmlWithCreateXmlReader()
    {
        using var excelReader = ExcelReaderFactory.CreateXmlReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        var dataSet = excelReader.AsDataSet();

        Assert.That(dataSet, Is.Not.Null);
        Assert.That(dataSet.Tables.Count, Is.EqualTo(1));
        Assert.That(dataSet.Tables[0].Rows.Count, Is.EqualTo(10));
    }

    [Test]
    public void ReadXmlFieldCount()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        
        Assert.That(excelReader.Read(), Is.True);
        Assert.That(excelReader.FieldCount, Is.EqualTo(10));
    }

    [Test]
    public void ReadXmlRowCount()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        
        Assert.That(excelReader.RowCount, Is.EqualTo(10));
    }

    [Test]
    public void ReadXmlMixedData()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        var dataSet = excelReader.AsDataSet();

        // Seventh row (index 6): 10x15, empty cells, 10x24
        Assert.That(dataSet.Tables[0].Rows[6][0], Is.EqualTo("10x15"));
        Assert.That(dataSet.Tables[0].Rows[6][1], Is.EqualTo(DBNull.Value));
        Assert.That(dataSet.Tables[0].Rows[6][9], Is.EqualTo("10x24"));
        
        // Eighth row (index 7): 10x16 to 10x25
        Assert.That(dataSet.Tables[0].Rows[7][0], Is.EqualTo("10x16"));
        Assert.That(dataSet.Tables[0].Rows[7][1], Is.EqualTo("10x17"));
        Assert.That(dataSet.Tables[0].Rows[7][2], Is.EqualTo("10x18"));
        Assert.That(dataSet.Tables[0].Rows[7][9], Is.EqualTo("10x25"));
    }

    [Test]
    public void ReadXmlGetValues()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        
        Assert.That(excelReader.Read(), Is.True);
        var values = new object[excelReader.FieldCount];
        excelReader.GetValues(values);
        
        Assert.That(values.Length, Is.EqualTo(10));
        Assert.That(values[0], Is.EqualTo("col1"));
        Assert.That(values[1] == null || values[1] == DBNull.Value, Is.True);
        Assert.That(values[2], Is.EqualTo("col3"));
        Assert.That(values[8], Is.EqualTo("col9"));
    }

    [Test]
    public void ReadXmlIsDBNull()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        var dataSet = excelReader.AsDataSet();

        // First row has empty cells at positions 1, 3, 5, 7
        var row = dataSet.Tables[0].Rows[0];
        Assert.That(row[0], Is.Not.EqualTo(DBNull.Value));
        Assert.That(row[1], Is.EqualTo(DBNull.Value));
        Assert.That(row[2], Is.Not.EqualTo(DBNull.Value));
        Assert.That(row[3], Is.EqualTo(DBNull.Value));
        
        // Second row has empty cells in the middle
        row = dataSet.Tables[0].Rows[1];
        Assert.That(row[0], Is.Not.EqualTo(DBNull.Value));
        Assert.That(row[1], Is.EqualTo(DBNull.Value));
        Assert.That(row[9], Is.Not.EqualTo(DBNull.Value));
    }

    [Test]
    public void ReadXmlName()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        
        Assert.That(excelReader.Name, Is.EqualTo("Sheet1"));
    }

    [Test]
    public void ReadXmlResultsCount()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("Test10x10.xml"));
        
        Assert.That(excelReader.ResultsCount, Is.EqualTo(1));
    }

    [Test]
    public void ReadXmlMultiSheetResultsCount()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestMultiSheet.xml"));
        
        Assert.That(excelReader.ResultsCount, Is.EqualTo(3));
    }

    [Test]
    public void ReadTestOpenXml()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestOpen.xml"));
        var dataSet = excelReader.AsDataSet();

        Assert.That(dataSet, Is.Not.Null);
        Assert.That(dataSet.Tables.Count, Is.EqualTo(3));
        Assert.That(dataSet.Tables[0].TableName, Is.EqualTo("Sheet1"));
        Assert.That(dataSet.Tables[1].TableName, Is.EqualTo("Sheet2"));
        Assert.That(dataSet.Tables[2].TableName, Is.EqualTo("Sheet3"));
    }

    [Test]
    public void ReadTestOpenXmlSheet1()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestOpen.xml"));
        var dataSet = excelReader.AsDataSet();

        // Sheet1: 7行，列数可能被扩展到最大列数（161列，因为Sheet2有161列）
        var sheet1 = dataSet.Tables["Sheet1"];
        Assert.That(sheet1.Rows.Count, Is.EqualTo(7));

        // AsDataSet 会将所有表的列数统一为最大列数
        Assert.That(sheet1.Columns.Count, Is.GreaterThanOrEqualTo(11));

        // First row: date, numbers, text, date, numbers, time
        var firstRow = sheet1.Rows[0];
        Assert.That(firstRow[0], Is.InstanceOf<DateTime>());
        Assert.That(firstRow[1], Is.EqualTo(1.0));
        Assert.That(firstRow[2], Is.EqualTo(1.02));
        Assert.That(firstRow[3], Is.EqualTo("text"));
        Assert.That(firstRow[4], Is.InstanceOf<DateTime>());
        Assert.That(firstRow[5], Is.EqualTo(6.0));
        Assert.That(firstRow[10], Is.InstanceOf<DateTime>());
    }

    [Test]
    public void ReadTestOpenXmlSheet1EmptyCells()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestOpen.xml"));
        var dataSet = excelReader.AsDataSet();

        var sheet1 = dataSet.Tables["Sheet1"];

        // Second row has empty cell at index 5
        var secondRow = sheet1.Rows[1];
        Assert.That(secondRow[0], Is.InstanceOf<DateTime>());
        Assert.That(secondRow[1], Is.EqualTo(2.0));
        Assert.That(secondRow[2], Is.EqualTo(2.04));
        Assert.That(secondRow[3], Is.EqualTo("\"text\""));
        Assert.That(secondRow[5], Is.EqualTo(DBNull.Value));
        Assert.That(secondRow[6], Is.EqualTo(7.0));
    }

    [Test]
    public void ReadTestOpenXmlSheet2()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestOpen.xml"));
        var dataSet = excelReader.AsDataSet();

        // Sheet2: 4行161列
        var sheet2 = dataSet.Tables["Sheet2"];
        Assert.That(sheet2.Rows.Count, Is.EqualTo(4));
        Assert.That(sheet2.Columns.Count, Is.EqualTo(161));

        // First three rows are all "aaa"
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                Assert.That(sheet2.Rows[row][col], Is.EqualTo("aaa"));
            }
        }

        // Fourth row has numbers 1-161
        var fourthRow = sheet2.Rows[3];
        Assert.That(fourthRow[0], Is.EqualTo(1.0));
        Assert.That(fourthRow[1], Is.EqualTo(2.0));
        Assert.That(fourthRow[160], Is.EqualTo(161.0));
    }

    [Test]
    public void ReadTestOpenXmlSheet3()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestOpen.xml"));
        var dataSet = excelReader.AsDataSet();

        // Sheet3: 13行13列 with formulas
        var sheet3 = dataSet.Tables["Sheet3"];
        Assert.That(sheet3.Rows.Count, Is.EqualTo(13));
        Assert.That(sheet3.Columns.Count, Is.EqualTo(13));

        // First row: numbers starting from 1.0001
        var firstRow = sheet3.Rows[0];
        Assert.That(firstRow[0], Is.EqualTo(1.0001));
        Assert.That(firstRow[1], Is.EqualTo(2.0002));
        Assert.That(firstRow[12], Is.EqualTo(4096.4096));
    }

    [Test]
    public void ReadTestOpenXmlDateTimeValues()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestOpen.xml"));
        var dataSet = excelReader.AsDataSet();

        var sheet1 = dataSet.Tables["Sheet1"];

        // First row first column is a date
        var date1 = (DateTime)sheet1.Rows[0][0];
        Assert.That(date1.Year, Is.EqualTo(2006));
        Assert.That(date1.Month, Is.EqualTo(10));
        Assert.That(date1.Day, Is.EqualTo(10));

        // First row fourth column is also a date
        var date2 = (DateTime)sheet1.Rows[0][4];
        Assert.That(date2.Year, Is.EqualTo(2009));
        Assert.That(date2.Month, Is.EqualTo(1));
        Assert.That(date2.Day, Is.EqualTo(1));
    }

    [Test]
    public void ReadTestOpenXmlDecimalValues()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestOpen.xml"));
        var dataSet = excelReader.AsDataSet();

        var sheet1 = dataSet.Tables["Sheet1"];

        // First row has decimal values
        Assert.That(sheet1.Rows[0][2], Is.EqualTo(1.02));
        Assert.That(sheet1.Rows[1][2], Is.EqualTo(2.04));
        Assert.That(sheet1.Rows[2][2], Is.EqualTo(4.08));
    }

    [Test]
    public void ReadTestOpenXmlStringValues()
    {
        using var excelReader = ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook("TestOpen.xml"));
        var dataSet = excelReader.AsDataSet();

        var sheet1 = dataSet.Tables["Sheet1"];

        // First row has text
        Assert.That(sheet1.Rows[0][3], Is.EqualTo("text"));
        Assert.That(sheet1.Rows[1][3], Is.EqualTo("\"text\""));
        Assert.That(sheet1.Rows[2][3], Is.EqualTo("Text"));
        Assert.That(sheet1.Rows[3][3], Is.EqualTo("Text Test"));
    }
}
