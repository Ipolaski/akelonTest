using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelWorkProject.Models;
using System.IO;

public class ActionsWithXml
{
    private readonly string _pathFile;

    public ActionsWithXml(string pathFile)
    {
        _pathFile = pathFile;
    }
    public static void Main()
    {

    }
    public List<ClientInfo>? GetClientInfoFromXml(string nameProduct)
    {
        var result = new List<ClientInfo>();

        var workBook = new XLWorkbook(_pathFile);

        var product = GetProductByName(workBook, nameProduct);
        if (product is null)
        {
            return result;
        }

        var requests = GetRequestsByProductCode(workBook, product.Code);
        foreach (var request in requests)
        {
            var client = GetClientByCode(workBook, request.CodeClient);
            var productPrice = GetProductByCode(workBook, request.CodeProduct);

            if (client is null)
                continue;

            result.Add(new ClientInfo(client.OrganizationName, client.Address, client.ContactPerson, request.Count, productPrice != null ? productPrice.Price : 0d, request.PostingDate));
        }

        return result;
    }
    
    public List<Client> GetAllClients()
    {
        var workBook = new XLWorkbook(_pathFile);

        return GetClients(workBook);
    }

    public Client? GetClientInfoByCode(int code)
    {
        var workBook = new XLWorkbook(_pathFile);

        return GetClientByCode(workBook, code);
    }

    public bool EditOrganizationNameClient(int code, string name)
    {
        var workBook = new XLWorkbook(_pathFile);
        var worksheet = workBook.Worksheet("Клиенты");

        IXLRow client;
        try
        {
            client = worksheet.CellsUsed(cell => cell.GetString() == code.ToString()).First().WorksheetRow();
        }
        catch
        {
            return false;
        }

        client.Cell(2).Value = name;
        try
        {
            workBook.Save();
        }
        catch
        {
            return false;
        }

        return true;
    }

    public bool EditAddressClient(int code, string address)
    {
        var workBook = new XLWorkbook(_pathFile);
        var worksheet = workBook.Worksheet("Клиенты");

        IXLRow client;
        try
        {
            client = worksheet.CellsUsed(cell => cell.GetString() == code.ToString()).First().WorksheetRow();
        }
        catch
        {
            return false;
        }

        client.Cell(3).Value = address;
        try
        {
            workBook.Save();
        }
        catch
        {
            return false;
        }

        return true;
    }

    public bool EditPersonInfoClient(int code, string person)
    {
        var workBook = new XLWorkbook(_pathFile);
        var worksheet = workBook.Worksheet("Клиенты");

        IXLRow client;
        try
        {
            client = worksheet.CellsUsed(cell => cell.GetString() == code.ToString()).First().WorksheetRow();
        }
        catch
        {
            return false;
        }

        client.Cell(4).Value = person;
        try
        {
            workBook.Save();
        }
        catch
        {
            return false;
        }

        return true;
    }

    public Client? GetGreatestOrdersClient(DateTime start, DateTime end)
    {
        var workBook = new XLWorkbook(_pathFile);

        var requests =  GetAllRequests(workBook);

        var sortClients = requests.Where(a => a.PostingDate >= start && a.PostingDate <= end);
        var clients = from client in sortClients
                  group client by client.CodeClient into g
                  select new { Code = g.Key, Count = g.Count() };

        int clientcode = clients.OrderByDescending(a => a.Count).FirstOrDefault() != null? clients.OrderByDescending(a => a.Count).FirstOrDefault()!.Code : 0;

        return GetClientByCode(workBook, clientcode);
    }

    private List<Client> GetClients(XLWorkbook workbook)
    {
        var worksheet = workbook.Worksheet("Клиенты");

        var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

        var result = new List<Client>();
        foreach (var row in rows)
        {
            int code = 0;
            try
            {
                code = (int)row.Cell(1).Value.GetNumber();
            }
            catch
            {
                continue;
            }
            result.Add(new Client(code, row.Cell(2).Value.GetText(), row.Cell(4).Value.GetText(), row.Cell(3).Value.GetText()));
        }

        return result;
    }

    private List<Request> GetAllRequests(XLWorkbook workbook)
    {
        var worksheet = workbook.Worksheet("Заявки");

        var requestRows = worksheet.RangeUsed().RowsUsed().Skip(1);

        var result = new List<Request>();
        foreach (var requestRow in requestRows)
        {
            int code, codeProduct, codeClient, numberRequest, count = 0;
            try
            {
                code = (int)requestRow.Cell(1).Value.GetNumber();
                codeProduct = (int)requestRow.Cell(2).Value.GetNumber();
                codeClient = (int)requestRow.Cell(3).Value.GetNumber();
                numberRequest = (int)requestRow.Cell(4).Value.GetNumber();
                count = (int)requestRow.Cell(5).Value.GetNumber();
            }
            catch
            {
                continue;
            }

            result.Add(new Request(code, codeProduct, codeClient, numberRequest, count, requestRow.Cell(6).Value.GetDateTime()));
        }

        return result;
    }

    private List<Request> GetRequestsByProductCode(XLWorkbook workbook, int productCode)
    {
        var worksheet = workbook.Worksheet("Заявки");

        IXLCells requestCells = worksheet.CellsUsed(cell => cell.GetString() == productCode.ToString() && cell.Address.ColumnNumber == 2);

        var result = new List<Request>();
        foreach (var cell in requestCells)
        {
            var requestRow = cell.WorksheetRow();

            int code, codeProduct, codeClient, numberRequest, count = 0;
            try
            {
                code = (int)requestRow.Cell(1).Value.GetNumber();
                codeProduct = (int)requestRow.Cell(2).Value.GetNumber();
                codeClient = (int)requestRow.Cell(3).Value.GetNumber();
                numberRequest = (int)requestRow.Cell(4).Value.GetNumber();
                count = (int)requestRow.Cell(5).Value.GetNumber();
            }
            catch
            {
                continue;
            }

            result.Add(new Request(code, codeProduct, codeClient, numberRequest, count, requestRow.Cell(6).Value.GetDateTime()));
        }

        return result;
    }

    private Product? GetProductByName(XLWorkbook workbook, string name)
    {
        var worksheet = workbook.Worksheet("Товары");

        IXLRow product;
        try
        {
            product = worksheet.CellsUsed(cell => cell.GetString() == name).First().WorksheetRow();
        }
        catch
        {
            return null;
        }

        int code = (int)product.Cell(1).Value.GetNumber();

        double price = product.Cell(4).Value.GetNumber();

        var result = new Product(code, product.Cell(2).Value.GetText(), product.Cell(3).Value.GetText(), price);

        return result;
    }

    private Client? GetClientByCode(XLWorkbook workbook, int code)
    {
        var worksheet = workbook.Worksheet("Клиенты");

        IXLRow client;
        try
        {
            client = worksheet.CellsUsed(cell => cell.GetString() == code.ToString()).First().WorksheetRow();
        }
        catch
        {
            return null;
        }

        var result = new Client(code, client.Cell(2).Value.GetText(), client.Cell(4).Value.GetText(), client.Cell(3).Value.GetText());

        return result;
    }

    private Product? GetProductByCode(XLWorkbook workbook, int code)
    {
        var worksheet = workbook.Worksheet("Товары");

        IXLRow product;
        try
        {
            product = worksheet.CellsUsed(cell => cell.GetString() == code.ToString()).First().WorksheetRow();
        }
        catch
        {
            return null;
        }

        double price = 0d;
        try
        {
            price = product.Cell(4).Value.GetNumber();
        }
        catch
        {
            return null;
        }

        var result = new Product(code, product.Cell(2).Value.GetText(), product.Cell(3).Value.GetText(), price);

        return result;
    }
}