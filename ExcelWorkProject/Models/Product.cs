namespace ExcelWorkProject.Models
{
    public class Product
    {
        public Product(int code, string name, string unitMeasurement, double price)
        {
            Code = code;
            Name = name;
            UnitMeasurement = unitMeasurement;
            Price = price;
        }

        public int Code { get; set; }
        public string Name { get; set; } = null!;
        public string UnitMeasurement { get; set; } = null!;
        public double Price { get; set; }
    }
}
