namespace ExcelWorkProject.Models
{
    public class ClientInfo
    {
        public ClientInfo( string organizationName, string adress, string contactPerson, int productQuantity, double priceForOne, DateTime orderDate )
        {
            OrganizationName = organizationName;
            Adress = adress;
            ContactPerson = contactPerson;
            ProductQuantity = productQuantity;
            PriceForOne = priceForOne;
            OrderDate = orderDate;
        }

        public string OrganizationName { get; set; }
        public string Adress {  get; set; }
        public string ContactPerson {  get; set; }
        public int ProductQuantity { get; set; }
        public double PriceForOne { get; set; }
        public DateTime OrderDate { get; set; }
    }
}