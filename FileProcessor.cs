public class FileProcessor
{
    private readonly string _inputDirectory;
    private readonly string _outputDirectory;
    private readonly IEnumerable<string> _possibleHeaders;
    private readonly string _splitAddressPattern;

    public FileProcessor(string inputDirectory, string outputDirectory, IEnumerable<string> possibleHeaders, string splitAddressPattern)
    {
        _inputDirectory = NormalizePath(inputDirectory);
        _outputDirectory = NormalizePath(outputDirectory);
        _possibleHeaders = possibleHeaders;
        _splitAddressPattern = splitAddressPattern;

        if (!Directory.Exists(_outputDirectory))
        {
            Console.WriteLine($"{DateTime.Now}: Output directory does not exist. Creating directory: {_outputDirectory}");
            Directory.CreateDirectory(_outputDirectory);
        }
    }

    public void ProcessFiles()
    {
        Console.WriteLine($"{DateTime.Now}: Starting file processing.");
        Console.WriteLine($"{DateTime.Now}: Input Directory = {_inputDirectory}");
        Console.WriteLine($"{DateTime.Now}: Output Directory = {_outputDirectory}");

        string[] files;

        try
        {
            files = Directory.GetFiles(_inputDirectory, "*.*");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"{DateTime.Now}: Error accessing input directory: {_inputDirectory}. Exception: {ex.Message}");
            return;
        }

        if (files.Length == 0)
        {
            Console.WriteLine($"{DateTime.Now}: No files found in the input directory: {_inputDirectory}");
            return;
        }

        Console.WriteLine($"{DateTime.Now}: Found {files.Length} file(s) to process.");

        foreach (string file in files)
        {
            string outputFilePath = Path.Combine(_outputDirectory, Path.GetFileName(file));
            outputFilePath = NormalizePath(outputFilePath);
            string normalizedFile = NormalizePath(file);   

            try
            {
                Console.WriteLine($"{DateTime.Now}: Copying file: {normalizedFile} to {outputFilePath}");
                File.Copy(file, outputFilePath, overwrite: true);

                if (file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"{DateTime.Now}: Processing .xlsx file: {outputFilePath}");
                    ExcelHelper.ProcessXlsx(outputFilePath, _possibleHeaders, _splitAddressPattern);
                }
                else if (file.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"{DateTime.Now}: Processing .xls file: {outputFilePath}");
                    ExcelHelper.ProcessXls(outputFilePath, _possibleHeaders, _splitAddressPattern);
                }
                else
                {
                    Console.WriteLine($"{DateTime.Now}: Skipping unsupported file type: {normalizedFile}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{DateTime.Now}: Error processing file: {normalizedFile}. Exception: {ex.Message}");
            }
        }

        Console.WriteLine($"{DateTime.Now}: File processing completed. Processed files are saved in: {_outputDirectory}");
    }

    private string NormalizePath(string path)
    {
        return Path.GetFullPath(path).Replace('\\', '/'); 
    }
}
