namespace WebApplication4.Services
{
    public class ConvertionResult<T> where T : new()
    {
        public List<T> OutputList { get; set; }

        public string ErrorMessage { get; set; }

        public ConvertionResult()
        {
            OutputList = new List<T>();
        }
    }
}
