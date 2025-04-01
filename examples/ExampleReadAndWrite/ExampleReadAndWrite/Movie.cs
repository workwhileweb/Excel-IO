using System.Collections.Generic;
using Excel.IO;

namespace ExampleReadAndWrite;

public class Movie : IExcelRow
{
    public static List<Movie> GetSampleMovies()
    {
        return
        [
            new Movie
            {
                Title = "The Shawshank Redemption", Director = "Frank Darabont", ReleaseYear = 1994, Genre = "Drama"
            },
            new Movie
            {
                Title = "The Godfather", Director = "Francis Ford Coppola", ReleaseYear = 1972, Genre = "Crime"
            },
            new Movie
            {
                Title = "The Dark Knight", Director = "Christopher Nolan", ReleaseYear = 2008, Genre = "Action"
            },
            new Movie { Title = "Pulp Fiction", Director = "Quentin Tarantino", ReleaseYear = 1994, Genre = "Crime" },
            new Movie
            {
                Title = "The Lord of the Rings: The Return of the King", Director = "Peter Jackson", ReleaseYear = 2003,
                Genre = "Fantasy"
            },
            new Movie { Title = "Forrest Gump", Director = "Robert Zemeckis", ReleaseYear = 1994, Genre = "Drama" },
            new Movie { Title = "Inception", Director = "Christopher Nolan", ReleaseYear = 2010, Genre = "Sci-Fi" },
            new Movie { Title = "Fight Club", Director = "David Fincher", ReleaseYear = 1999, Genre = "Drama" },
            new Movie
            {
                Title = "The Matrix", Director = "Lana Wachowski, Lilly Wachowski", ReleaseYear = 1999, Genre = "Sci-Fi"
            },
            new Movie { Title = "Goodfellas", Director = "Martin Scorsese", ReleaseYear = 1990, Genre = "Crime" }
        ];
    }
    public string SheetName => "Movies Sheet";

    public string Title { get; set; }

    public string Director { get; set; }

    public int ReleaseYear { get; set; }

    public string Genre { get; set; }
}