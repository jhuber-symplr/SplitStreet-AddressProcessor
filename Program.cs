using Microsoft.Extensions.Configuration;
using System.Diagnostics;

class Program
{
    static void Main()
    {
        Stopwatch stopwatch = Stopwatch.StartNew();

        Console.WriteLine($"{DateTime.Now}: Starting the application.");

        var configuration = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        string inputDirectory = configuration["FileProcessingSettings:InputDirectory"];
        string outputDirectory = configuration["FileProcessingSettings:OutputDirectory"];
        var possibleHeaders = configuration.GetSection("FileProcessingSettings:PossibleHeaders").Get<List<string>>();
        string splitAddressPattern = configuration["FileProcessingSettings:SplitAddressPattern"];

        if (string.IsNullOrWhiteSpace(splitAddressPattern))
        {
            Console.WriteLine($"{DateTime.Now}: Split address pattern is not configured in appsettings.json.");
            return;
        }

        if (string.IsNullOrWhiteSpace(inputDirectory) || string.IsNullOrWhiteSpace(outputDirectory))
        {
            Console.WriteLine($"{DateTime.Now}: Input or output directory is not configured in appsettings.json.");
            return;
        }

        Console.WriteLine($"{DateTime.Now}: Loaded configuration.");
        Console.WriteLine($"Input Directory = {inputDirectory}");
        Console.WriteLine($"Output Directory = {outputDirectory}");
        Console.WriteLine($"Possible Headers = {string.Join(", ", possibleHeaders)}");

        var processor = new FileProcessor(inputDirectory, outputDirectory, possibleHeaders, splitAddressPattern);

        try
        {
            processor.ProcessFiles();
            Console.WriteLine($"{DateTime.Now}: File processing completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"{DateTime.Now}: Error occurred during file processing. Exception: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }

        stopwatch.Stop();
        Console.WriteLine($"{DateTime.Now}: Application finished in {stopwatch.ElapsedMilliseconds} ms.");
    }
}
