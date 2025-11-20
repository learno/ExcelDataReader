using System.Globalization;
using System.Text;
using System.Xml;

namespace ExcelDataReader.Core.XmlFormat;

internal sealed class XmlWorksheet : IWorksheet
{
    private readonly Stream _stream;
    private readonly string _name;
    private int _fieldCount;
    private int _rowCount;
    private List<Column> _columnWidths;
    private CellRange[] _mergeCells;

    public XmlWorksheet(Stream stream, string name)
    {
        _stream = stream;
        _name = name;
        AnalyzeWorksheet();
    }

    public string Name => _name;

    public string CodeName => null;

    public string VisibleState => null;

    public HeaderFooter HeaderFooter => null;

    public int FieldCount => _fieldCount;

    public int RowCount => _rowCount;

    public CellRange[] MergeCells => _mergeCells;

    public List<Column> ColumnWidths => _columnWidths;

    private void AnalyzeWorksheet()
    {
        _stream.Seek(0, SeekOrigin.Begin);
        using var xmlReader = XmlReader.Create(_stream, new XmlReaderSettings { IgnoreWhitespace = true, IgnoreComments = true });

        var maxColumns = 0;
        var rowCount = 0;
        var inTargetWorksheet = false;
        var worksheetDepth = 0;
        var cellRanges = new List<CellRange>();

        while (xmlReader.Read())
        {
            if (xmlReader.NodeType == XmlNodeType.Element)
            {
                var localName = xmlReader.LocalName;
                
                if (localName == "Worksheet" || localName == "ss:Worksheet")
                {
                    var name = xmlReader.GetAttribute("Name") ?? xmlReader.GetAttribute("ss:Name");
                    if (name == _name || (!inTargetWorksheet && name == null))
                    {
                        inTargetWorksheet = true;
                        worksheetDepth = xmlReader.Depth;
                    }
                }
                else if (inTargetWorksheet)
                {
                    if (localName == "Row" || localName == "ss:Row")
                    {
                        rowCount++;
                        var cellCount = 0;
                        var rowDepth = xmlReader.Depth;
                        var currentRowIndex = rowCount - 1; // 0-based row index
                        var currentColumnIndex = 0;
                        
                        while (xmlReader.Read() && xmlReader.Depth > rowDepth)
                        {
                            if (xmlReader.NodeType == XmlNodeType.Element)
                            {
                                if (xmlReader.LocalName == "Cell" || xmlReader.LocalName == "ss:Cell")
                                {
                                    // 检查Index属性
                                    var indexAttr = xmlReader.GetAttribute("Index") ?? xmlReader.GetAttribute("ss:Index");
                                    if (!string.IsNullOrEmpty(indexAttr) && int.TryParse(indexAttr, out var index))
                                    {
                                        currentColumnIndex = index - 1; // Index是1-based
                                    }

                                    // 检查MergeAcross属性（合并列）
                                    var mergeAcrossAttr = xmlReader.GetAttribute("MergeAcross") ?? xmlReader.GetAttribute("ss:MergeAcross");
                                    var mergeAcross = 0;
                                    if (!string.IsNullOrEmpty(mergeAcrossAttr) && int.TryParse(mergeAcrossAttr, out var mergeCount))
                                    {
                                        mergeAcross = mergeCount;
                                    }

                                    // 检查MergeDown属性（合并行）
                                    var mergeDownAttr = xmlReader.GetAttribute("MergeDown") ?? xmlReader.GetAttribute("ss:MergeDown");
                                    var mergeDown = 0;
                                    if (!string.IsNullOrEmpty(mergeDownAttr) && int.TryParse(mergeDownAttr, out var mergeDownCount))
                                    {
                                        mergeDown = mergeDownCount;
                                    }

                                    // 如果有合并，记录合并范围
                                    if (mergeAcross > 0 || mergeDown > 0)
                                    {
                                        var fromColumn = currentColumnIndex;
                                        var fromRow = currentRowIndex;
                                        var toColumn = currentColumnIndex + mergeAcross;
                                        var toRow = currentRowIndex + mergeDown;
                                        cellRanges.Add(new CellRange(fromColumn, fromRow, toColumn, toRow));
                                    }

                                    cellCount++;
                                    cellCount = Math.Max(cellCount, currentColumnIndex + 1);
                                    
                                    // 更新列索引（包括合并的列）
                                    currentColumnIndex += 1 + mergeAcross;
                                }
                            }
                        }
                        
                        maxColumns = Math.Max(maxColumns, cellCount);
                    }
                    else if (localName == "Table" || localName == "ss:Table")
                    {
                        // 在Table级别，继续处理
                    }
                }
            }
            else if (xmlReader.NodeType == XmlNodeType.EndElement && inTargetWorksheet)
            {
                if (xmlReader.LocalName == "Worksheet" || xmlReader.LocalName == "ss:Worksheet")
                {
                    if (xmlReader.Depth == worksheetDepth - 1)
                    {
                        break;
                    }
                }
            }
        }

        _fieldCount = maxColumns;
        _rowCount = rowCount;
        _columnWidths = null;
        _mergeCells = [.. cellRanges];
    }

    public IEnumerable<Row> ReadRows()
    {
        _stream.Seek(0, SeekOrigin.Begin);
        using var xmlReader = XmlReader.Create(_stream, new XmlReaderSettings { IgnoreWhitespace = true, IgnoreComments = true });

        var inTargetWorksheet = false;
        var worksheetDepth = 0;
        var rowIndex = 0;

        // 移动到第一个元素
        if (xmlReader.Read() && xmlReader.NodeType == XmlNodeType.XmlDeclaration)
        {
            xmlReader.Read(); // 跳过 XML 声明
        }

        while (xmlReader.Read())
        {
            if (xmlReader.NodeType == XmlNodeType.Element)
            {
                var localName = xmlReader.LocalName;
                
                if (localName == "Worksheet" || localName == "ss:Worksheet")
                {
                    var name = xmlReader.GetAttribute("Name") ?? xmlReader.GetAttribute("ss:Name");
                    if (name == _name || (!inTargetWorksheet && name == null))
                    {
                        inTargetWorksheet = true;
                        worksheetDepth = xmlReader.Depth;
                    }
                    else if (inTargetWorksheet)
                    {
                        // 如果已经在目标工作表中，但遇到了另一个工作表，说明已经读取完毕
                        break;
                    }
                }
                else if (inTargetWorksheet && (localName == "Row" || localName == "ss:Row"))
                {
                    var row = ReadRow(xmlReader, rowIndex);
                    if (row != null)
                    {
                        yield return row;
                        rowIndex++;
                    }
                }
            }
            else if (xmlReader.NodeType == XmlNodeType.EndElement && inTargetWorksheet)
            {
                if (xmlReader.LocalName == "Worksheet" || xmlReader.LocalName == "ss:Worksheet")
                {
                    // 当遇到目标工作表的结束标签时，停止读取
                    if (xmlReader.Depth == worksheetDepth)
                    {
                        break;
                    }
                }
            }
        }
    }

    private Row ReadRow(XmlReader xmlReader, int rowIndex)
    {
        var cells = new List<Cell>();
        var rowDepth = xmlReader.Depth;
        var currentColumnIndex = 0;
        double? rowHeight = null;

        // 读取Row属性
        var heightAttr = xmlReader.GetAttribute("Height") ?? xmlReader.GetAttribute("ss:Height");
        if (!string.IsNullOrEmpty(heightAttr) && double.TryParse(heightAttr, NumberStyles.Float, CultureInfo.InvariantCulture, out var height))
        {
            rowHeight = height;
        }

        // 移动到Row的第一个子节点
        if (xmlReader.Read() && xmlReader.Depth > rowDepth)
        {
            do
            {
                if (xmlReader.NodeType == XmlNodeType.Element)
                {
                    if (xmlReader.LocalName == "Cell" || xmlReader.LocalName == "ss:Cell")
                    {
                        var cell = ReadCell(xmlReader, ref currentColumnIndex);
                        // 总是添加单元格，即使值为null（表示空单元格）
                        // 注意：不再为合并列创建空单元格，与 xlsx 格式保持一致
                        cells.Add(cell);
                    }
                }
            }
            while (xmlReader.Depth > rowDepth && xmlReader.Read());
        }

        return new Row(rowIndex, rowHeight ?? 12.75, cells);
    }

    private Cell ReadCell(XmlReader xmlReader, ref int currentColumnIndex)
    {
        // 检查Index属性
        var indexAttr = xmlReader.GetAttribute("Index") ?? xmlReader.GetAttribute("ss:Index");
        if (!string.IsNullOrEmpty(indexAttr) && int.TryParse(indexAttr, out var index))
        {
            currentColumnIndex = index - 1; // Index是1-based
        }

        // 检查MergeAcross属性（合并列）
        var mergeAcrossAttr = xmlReader.GetAttribute("MergeAcross") ?? xmlReader.GetAttribute("ss:MergeAcross");
        var mergeAcross = 0;
        if (!string.IsNullOrEmpty(mergeAcrossAttr) && int.TryParse(mergeAcrossAttr, out var mergeCount))
        {
            mergeAcross = mergeCount;
        }

        var columnIndex = currentColumnIndex;
        object value = null;
        var cellDepth = xmlReader.Depth;
        var isEmptyElement = xmlReader.IsEmptyElement;

        // 如果不是空元素，读取Cell内容
        if (!isEmptyElement)
        {
            // 移动到Cell的第一个子节点
            if (xmlReader.Read() && xmlReader.Depth > cellDepth)
            {
                do
                {
                    if (xmlReader.NodeType == XmlNodeType.Element)
                    {
                        if (xmlReader.LocalName == "Data" || xmlReader.LocalName == "ss:Data")
                        {
                            var typeAttr = xmlReader.GetAttribute("Type") ?? xmlReader.GetAttribute("ss:Type");
                            string innerText;
                            
                            // 如果元素是空元素或只有文本内容，使用 ReadElementContentAsString
                            // 否则使用 ReadInnerXml 来读取内容
                            if (xmlReader.IsEmptyElement)
                            {
                                innerText = string.Empty;
                                xmlReader.Read(); // 移动到下一个节点
                            }
                            else
                            {
                                // 检查是否有子元素
                                var depth = xmlReader.Depth;
                                xmlReader.Read(); // 移动到第一个子节点
                                
                                if (xmlReader.NodeType == XmlNodeType.Text || xmlReader.NodeType == XmlNodeType.CDATA)
                                {
                                    innerText = xmlReader.Value;
                                    // 移动到 Data 元素的结束标签
                                    while (xmlReader.Depth > depth && xmlReader.Read()) { }
                                }
                                else
                                {
                                    // 有子元素，使用 ReadInnerXml
                                    xmlReader.MoveToElement(); // 回到 Data 元素
                                    innerText = xmlReader.ReadInnerXml();
                                }
                            }
                            
                            value = ParseCellValue(innerText, typeAttr);
                        }
                        else if (xmlReader.LocalName == "Comment" || xmlReader.LocalName == "ss:Comment")
                        {
                            // 跳过注释元素
                            xmlReader.Skip();
                        }
                        else if (xmlReader.LocalName == "NamedCell" || xmlReader.LocalName == "ss:NamedCell")
                        {
                            // 跳过命名单元格
                            xmlReader.Skip();
                        }
                        else
                        {
                            // 跳过其他未知元素
                            xmlReader.Skip();
                        }
                    }
                    else if (xmlReader.NodeType == XmlNodeType.Text)
                    {
                        // 如果有文本内容但没有Data元素
                        if (value == null && !string.IsNullOrWhiteSpace(xmlReader.Value))
                        {
                            value = xmlReader.Value.Trim();
                        }
                    }
                }
                while (xmlReader.Depth > cellDepth && xmlReader.Read());
            }
        }

        // 如果xmlReader还在Cell结束标签上，需要读取下一个元素
        // 但不要在这里读取，让ReadRow的循环来处理

        // 更新列索引
        // 如果有合并列，需要跳过合并的列数
        currentColumnIndex += 1 + mergeAcross;
        
        // 即使值为null，也要返回Cell对象（表示空单元格）
        return new Cell(columnIndex, value, ExtendedFormat.Zero, null);
    }

    private object ParseCellValue(string text, string type)
    {
        if (string.IsNullOrEmpty(text))
        {
            return null;
        }

        if (string.IsNullOrEmpty(type))
        {
            // 如果没有指定类型，尝试自动推断
            return text;
        }

        switch (type.ToLowerInvariant())
        {
            case "string":
            case "str":
                return text;
            
            case "number":
                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                {
                    return number;
                }
                return text;
            
            case "boolean":
            case "bool":
                if (bool.TryParse(text, out var boolValue))
                {
                    return boolValue;
                }
                // Excel XML有时使用1/0表示布尔值
                if (text == "1" || text.ToLowerInvariant() == "true")
                {
                    return true;
                }
                if (text == "0" || text.ToLowerInvariant() == "false")
                {
                    return false;
                }
                return text;
            
            case "datetime":
            case "date":
                // 尝试解析 ISO 8601 格式 (2009-05-01T00:00:00 或 2009-05-01T00:00:00.000)
                if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateTime))
                {
                    return dateTime;
                }
                // 尝试解析其他常见格式
                var formats = new[]
                {
                    "yyyy-MM-ddTHH:mm:ss.fff",
                    "yyyy-MM-ddTHH:mm:ss",
                    "yyyy-MM-ddTHH:mm:ss.ff",
                    "yyyy-MM-ddTHH:mm:ss.f",
                    "yyyy-MM-dd",
                    "yyyy/MM/dd",
                    "MM/dd/yyyy"
                };
                foreach (var format in formats)
                {
                    if (DateTime.TryParseExact(text, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime))
                    {
                        return dateTime;
                    }
                }
                return text;
            
            case "error":
                return text;
            
            default:
                return text;
        }
    }
}

