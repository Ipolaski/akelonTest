namespace ExcelWorkProject.Models
{
    public class Client
    {
        public Client( int code, string organizationName, string contactPerson, string address)
        {
            Code = code;
            OrganizationName = organizationName;
            ContactPerson = contactPerson;
            Address = address;
        }

        public int Code { get; set; }
        public string OrganizationName { get; set; }
        public string ContactPerson { get; set; }
        public string Address { get; set; }
    }
}