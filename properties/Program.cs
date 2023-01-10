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
testList.Add(smolTest1);
testList.Add(smolTest2);
testList.Add(smolTest3);


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
    var ws = workBook.Worksheets.Worksheet(p.Name[..^6]); // Get the one just given

    var type = p.PropertyType.GetGenericArguments().Single();
    var classElementToGetColumns = Activator.CreateInstance(type); // if empty collection im not screwed.
    var property = p!.GetValue(propertiesBiggieData, null);

    WorkSheedDataFill(ws, property, classElementToGetColumns);

}



workBook.Properties.Title = "xx";
workBook.SaveAs(@$"C:\Users\lbund\Documents\{workBook.Properties.Title}.xlsx");

void WorkSheedDataFill(IXLWorksheet ws, object? objectData, object? obcjetColumns)
{
    var collection = (IList)objectData; // dane kolekcji

    // class with: columnName, 

    // kolumny only + puste dane
    var columnsProperties = obcjetColumns.GetType().GetProperties();
    var columnNames = columnsProperties.Select(p => p.Name).ToList();


    var oclumnData = new Dictionary<string, int>(); // name of the column and column number

    var column = 1;

    foreach (var columnName in columnNames)
    {
        Console.WriteLine($"(1, {column})" + " " + columnName);
        ws.Cell(1, column).Value = columnName;
        oclumnData.Add(columnName, column);
        column++;
    }

    // get type of class we got
    if (collection.Count > 0) 
    {
        var rowPerItem = new Dictionary<string, int>();

        foreach (var item in collection)
        {
            var properties = item.GetType().GetProperties();

            // save rows for each property too
            var row1 = 0;
            foreach (var a in properties)
            {
                var name = a.Name;
                if (!rowPerItem.ContainsKey(name))
                {
                    row1 = 2;
                    rowPerItem.Add(name, row1);
                }
                else
                {
                    // get row by name and +1
                    row1 = rowPerItem[name];
                    row1++;
                }


                //rowPerItem[name] = row1;
                var column2 = oclumnData[name];

                var rowValue = a.GetValue(item) != null ? a.GetValue(item)?.ToString() : string.Empty;
                Console.WriteLine(rowValue + " (" + row1 + " , " + column2 + ")");
                ws.Cell(row1, column2).Value = rowValue;
            }
        }
    }

    ws.Columns().AdjustToContents();
}