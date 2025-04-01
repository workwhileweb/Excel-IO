using System;
using System.Collections.Generic;
using System.Text.Json;
using Excel.IO;

namespace ExampleReadAndWrite;

internal class Program
{
    private static void Main()
    {
        var excelConverter = new ExcelConverter();
        
        excelConverter.Write(Movie.GetSampleMovies(), @"..\..\..\movie.xlsx","movie");
        foreach (var movie in excelConverter.Read<Movie>(@"..\..\..\movie.xlsx","movie"))
        {
            Console.WriteLine(JsonSerializer.Serialize(movie));
        }

        //Read Example            
        var people = excelConverter.Read<Person>(@"..\..\..\people.xlsx");

        foreach (var person in people)
        {
            Console.WriteLine($"{person.EyeColour} : {person.Age} : {person.Height}");
        }

        //Write Example
        var peopleToWrite = new List<Person>();

        for (var i = 0; i < 10; i++)
        {
            peopleToWrite.Add(new Person
            {
                EyeColour = Guid.NewGuid().ToString(),
                Age = new Random().Next(1, 100),
                Height = new Random().Next(100, 200)
            });
        }
            
        excelConverter.Write(peopleToWrite, @"..\..\..\newPeople.xlsx");
    }
}