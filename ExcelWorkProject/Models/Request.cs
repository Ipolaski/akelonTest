namespace ExcelWorkProject.Models
{
    public class Request
    {
        public Request(int code, int codeProduct, int codeClient, int numberRequest, int count, DateTime postringDate)
        {
            Code = code;
            CodeProduct = codeProduct;
            CodeClient = codeClient;
            NumberRequest = numberRequest;
            Count = count;
            PostingDate = postringDate;
        }
        public int Code { get; set; }
        public int CodeProduct { get; set; }
        public int CodeClient { get; set; }
        public int NumberRequest { get; set; }
        public int Count { get; set; }
        public DateTime PostingDate { get; set; }
    }
}
