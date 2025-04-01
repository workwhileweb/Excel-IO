using Excel.IO;

namespace ExampleReadAndWrite;

public class Movie : IExcelRow
{
    public string SheetName => "Movies Sheet";

    public string Title { get; set; }

    public string Director { get; set; }

    public int ReleaseYear { get; set; }

    public string Genre { get; set; }
}
