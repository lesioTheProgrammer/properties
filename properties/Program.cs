using ClosedXML.Excel;
using properties;
using System.Collections;

var smol1 = new Smol1Report()
{
    Name1 = "a",
    Name2 = "b"

};

var smol2 = new Smol1Report()
{
    Name1 = "a1",
    Name2 = "b1"
};

var smolTest1 = new Test1()
{
    Name3 = "buda",
    Name4 = "psina",
    Double1 = 1

};

var smolTest2 = new Test1()
{
    Name3 = "buda",
    Name4 = "psina",
    Double1 = 1

};

var smolTest3 = new Test1()
{
    Name3 = "buda",
    Name4 = "psina"
};

var smol3 = new Smol2Report()
{
    Id1 = 8,
    Id2 = 88

};

var smol4 = new Smol2Report()
{
    Id1 = 6,
    Id2 = 666
};

var listSmol1 = new List<Smol1Report>();
listSmol1.Add(smol1);
listSmol1.Add(smol2);

var listSmol2 = new List<Smol2Report>();
listSmol2.Add(smol3);
listSmol2.Add(smol4);


var testList = new List<Test1>();


var propertiesBiggieData = new PropertiesBiggie()
{
    Smol1Report = listSmol1,
    Smol2Report = listSmol2,
    Test1Report = testList
};




var workBook = new XLWorkbook();
var properties = propertiesBiggieData.GetType().GetProperties();
// Create sheets


foreach (var p in properties)
{
    workBook.Worksheets.Add(p.Name[..^6]); // Adds sheet with given name

    var ws = workBook.Worksheets.Worksheet(p.Name[..^6]);
    var data = (IEnumerable<IReportData>)p!.GetValue(propertiesBiggieData, null);
    PoulateWorkSheet(ws, data!);

}



workBook.Properties.Title = "xx";
workBook.SaveAs(@$"C:\Users\lbund\Documents\{workBook.Properties.Title}.xlsx");

 void PoulateWorkSheet(IXLWorksheet ws, IEnumerable<IReportData> data)
{
    var listType = data.GetType();
    var itemType = listType.GetGenericArguments()[0];
    var itemProperties = itemType.GetProperties();
    var columnNames = itemProperties.Select(p => p.Name);
    var column = 1;
    var row = 1;

    //populate first raw with column names
    foreach (var columnName in columnNames)
    {
        ws.Cell(row, column++).Value = columnName;
    }

    foreach (var item in data)
    {
        row++;
        column = 1;
        foreach (var property in itemProperties)
        {
            var propertyValue = property.GetValue(item);
            ws.Cell(row, column++).Value = propertyValue == null ? string.Empty : propertyValue.ToString();
        }
    }

    ws.Columns().AdjustToContents();
}