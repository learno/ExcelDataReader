using System.Xml;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core.XmlFormat;

internal sealed class XmlWorkbook : IWorkbook<XmlWorksheet>
{
    private readonly Stream _stream;
    private List<string> _worksheetNames;

    public XmlWorkbook(Stream stream)
    {
        _stream = stream;
        _worksheetNames = new List<string>();
        ScanWorksheets();
    }

    public int ResultsCount => _worksheetNames.Count;

    public int ActiveSheet => 0;

    private void ScanWorksheets()
    {
        _stream.Seek(0, SeekOrigin.Begin);
        using var xmlReader = XmlReader.Create(_stream, new XmlReaderSettings { IgnoreWhitespace = true, IgnoreComments = true });

        // 查找Workbook根元素
        while (xmlReader.Read())
        {
            if (xmlReader.NodeType == XmlNodeType.Element)
            {
                var localName = xmlReader.LocalName;
                if (localName == "Workbook" || localName == "ss:Workbook")
                {
                    // 扫描所有Worksheet名称
                    while (xmlReader.Read())
                    {
                        if (xmlReader.NodeType == XmlNodeType.Element)
                        {
                            if (xmlReader.LocalName == "Worksheet" || xmlReader.LocalName == "ss:Worksheet")
                            {
                                var name = xmlReader.GetAttribute("Name") ?? xmlReader.GetAttribute("ss:Name");
                                if (string.IsNullOrEmpty(name))
                                {
                                    // 如果没有名称，使用默认名称
                                    name = $"Sheet{_worksheetNames.Count + 1}";
                                }
                                _worksheetNames.Add(name);
                            }
                        }
                        else if (xmlReader.NodeType == XmlNodeType.EndElement && 
                                 (xmlReader.LocalName == "Workbook" || xmlReader.LocalName == "ss:Workbook"))
                        {
                            break;
                        }
                    }
                    break;
                }
            }
        }

        // 如果没有找到任何worksheet，至少创建一个默认的
        if (_worksheetNames.Count == 0)
        {
            _worksheetNames.Add("Sheet1");
        }
    }

    public IEnumerable<XmlWorksheet> ReadWorksheets()
    {
        foreach (var name in _worksheetNames)
        {
            yield return new XmlWorksheet(_stream, name);
        }
    }

    public NumberFormatString GetNumberFormatString(int index)
    {
        return null;
    }
}

