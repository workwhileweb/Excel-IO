using Excel.IO;

namespace ExampleReadAndWrite;

public class Person : IExcelRow
{
    public string SheetName => "People Sheet";

    public string EyeColour { get; set; }

    public int Age { get; set; }

    public int Height { get; set; }
}