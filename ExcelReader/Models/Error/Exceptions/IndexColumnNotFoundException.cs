namespace ExcelReader.ExcelReader.Models.Error.Exceptions;

public class IndexColumnNotFoundException : Exception
{
    public IndexColumnNotFoundException()
    {
    }

    public IndexColumnNotFoundException(string message)
        : base(message)
    {
    }

    public IndexColumnNotFoundException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}