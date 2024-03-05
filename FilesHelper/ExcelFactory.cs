namespace FilesHelper
{
    public class ExcelFactory
    {
        public static ExcelFactory Instance { get { return new ExcelFactory(); } }

        public ExcelFactory() { }

        public FileStream GenerateEmpty()
        {
            return null;
        }

        public FileStream GenerateEmpty(List<string> columns)
        {
            return null;
        }
    }
}