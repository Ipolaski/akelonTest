using ExcelWorkProject.Models;

namespace MainApplication
{
    public class MenuDialog
    {
        public string GetDirectoryDialog()
        {
            string directory = string.Empty;
            string value = string.Empty;

            while (directory != "exit" && directory == string.Empty)
            {
                Console.WriteLine( "" +
                "Введите путь, включая имя и расширение файла.\n" +
                "Введите \"exit\" для выхода" );
                directory = Console.ReadLine()!;
                Console.WriteLine( directory );

                if (directory == "exit")
                {
                    value = directory;
                }
                else
                {
                    if (!File.Exists( directory ))
                    {
                        Console.WriteLine(
                            "Указанный файл не найден. \n" +
                            "Проверьте путь и имя файла. \n" +
                            "Нажмите Enter, чтобы указать путь заново." );
                        Console.ReadLine();
                        Console.Clear();

                        value = string.Empty;
                    }
                    else
                    {
                        value = directory;
                    }
                }
            }

            return value;
        }
        public string GetWorkWithFileDialog( string fileDirectory )
        {
            string userInput = string.Empty;
            ActionsWithXml xmlAction = new(fileDirectory);
            while (userInput != "exit" && fileDirectory != "exit")
            {
                Console.WriteLine(
                    "Введите цифру, соответствующую необходимому действию\n" +
                    "1 - вывести информацию о клиентах\n" +
                    "2 - Запрос на изменение контактного лица клиента\n" +
                    "3 - Запрос на определение золотого клиента\n" +
                    "Введите \"exit\" для выхода" );
                userInput = Console.ReadLine()!;

                switch (userInput)
                {
                    case "1":
                        Console.WriteLine("Введите наименование товара: ");
                        string product = Console.ReadLine()!;
                        PrintDisplayClientInfoProduct(xmlAction.GetClientInfoFromXml( product )!);
                        break;
                    case "2":
                        ChangeClientDialog(xmlAction);
                        break;
                    case "3":
                        GoldClientDialog(xmlAction);
                        break;
                    case "exit":
                        Console.Clear();
                        break;
                    default:
                        Console.WriteLine(
                            "Проверьте правильность ввода. Необходимо ввести только одно число.\n" +
                            "Нажмите Enter для продолжения" );
                        Console.ReadLine();
                        Console.Clear();
                        break;
                }
            }

            return userInput;
        }

        private void ChangeClientDialog(ActionsWithXml actionsWithXml)
        {
            string userInput = "";
            while (userInput != "exit")
            {
                Console.WriteLine("Введите код клиента для изменения информации\nall - вывод всех клиентов\nexit - назад в меню");
                userInput = Console.ReadLine()!;

                switch (userInput)
                {
                    case "all":
                        PrintDisplayAllClients(actionsWithXml.GetAllClients());
                        break;
                    case "exit":
                        Console.Clear();
                        break;
                    default:
                        {
                            int codePerson = 0;
                            int.TryParse(userInput, out codePerson);

                            if(actionsWithXml.GetClientInfoByCode(codePerson) == null)
                            {
                                Console.WriteLine("Такого пользователя не существует");
                                
                                break;
                            }
                            Console.WriteLine("Что вы хотите изменить?\nнаименование организации - 1\nАдрес - 2\nКонтактное лицо (ФИО) - 3");
                            userInput = Console.ReadLine()!;
                            switch (userInput)
                            {
                                case "1":
                                    Console.WriteLine("Введите новое наименование организации:");
                                    string nameOrg = Console.ReadLine()!;

                                    if(actionsWithXml.EditOrganizationNameClient(codePerson, nameOrg))
                                    {
                                        Console.WriteLine("Изменения успешо сохранены");
                                    }
                                    else
                                        Console.WriteLine("Не удалось сохранить изменения");
                                    break;

                                case "2":
                                    Console.WriteLine("Введите новый адрес:");
                                    string newAddress = Console.ReadLine()!;

                                    if(actionsWithXml.EditAddressClient(codePerson, newAddress))
                                    {
                                        Console.WriteLine("Изменения успешо сохранены");
                                    }
                                    else
                                        Console.WriteLine("Не удалось сохранить изменения");
                                    break;
                                    
                                case "3":
                                    Console.WriteLine("Введите новые контактные данные:");
                                    string newPersonInfo = Console.ReadLine()!;

                                    if (actionsWithXml.EditPersonInfoClient(codePerson, newPersonInfo))
                                    {
                                        Console.WriteLine("Изменения успешо сохранены");
                                    }
                                    else
                                        Console.WriteLine("Не удалось сохранить изменения");
                                    break;

                                default:
                                    Console.WriteLine("Команда не распознана");
                                    break;
                            }
                            break;
                        }
                }
            }
        }

        private void GoldClientDialog(ActionsWithXml actionsWithXml)
        {
            string userInput = "";
            while (userInput != "exit")
            {
                Console.WriteLine("1 - вывод клиента с наибольши количеством заказов\nexit - назад в меню");
                userInput = Console.ReadLine()!;

                switch (userInput)
                {
                    case "1":
                        (DateTime start, DateTime end) =  GetDateRange();
                        var displayResult = new List<Client>();
                        
                        var result = actionsWithXml.GetGreatestOrdersClient(start, end);
                        if (result == null)
                        {
                            Console.WriteLine("Ничего не найдено");
                            break;
                        }
                        
                        displayResult.Add(result);  
                        PrintDisplayAllClients(displayResult);
                        break;
                    case "exit":
                        break;
                    default:
                        Console.WriteLine("Команда не распознана");
                        break;
                }
            }
        }

        private (DateTime, DateTime) GetDateRange()
        {
            DateTime start = DateTime.Now;
            DateTime end = DateTime.Now;

            bool isConverted = false;
            
            int yearInt = 0;
            int monthInt = 0;
            while (!isConverted)
            {
                Console.WriteLine("Введите год:");
                string year = Console.ReadLine()!;

                
                if (int.TryParse(year, out yearInt))
                {
                    if (yearInt>1900 && yearInt < 2999)
                    {
                        isConverted = true;
                    }
                    else
                        Console.WriteLine("Введён не верный год");
                }
                else
                {
                    Console.WriteLine("Введён не верный год");
                }
            }

            isConverted = false;
            while (!isConverted)
            {
                Console.WriteLine("Введите месяц:");
                string month = Console.ReadLine()!;


                if (int.TryParse(month, out monthInt))
                {
                    if (monthInt > 0 && monthInt < 13)
                    {
                        isConverted = true;
                    }
                    else
                        Console.WriteLine("Введён не верный месяц");
                }
                else
                    Console.WriteLine("Введён не верный месяц");
            }

            start = new DateTime(yearInt, monthInt, 1);
            end = new DateTime(yearInt, monthInt, 1);
            end = end.AddMonths(1);

            return (start, end);
        }

        private void PrintDisplayClientInfoProduct(List<ClientInfo> clients)
        {
            Console.WriteLine($"Название организации\tАдрес\tФИО\tКоличество\tЦена\tДата размещения");
            foreach (var clientInfo in clients)
            {
                Console.WriteLine($"{clientInfo.OrganizationName}\t{clientInfo.Adress}\t{clientInfo.ContactPerson}\t{clientInfo.ProductQuantity}\t{clientInfo.PriceForOne}\t{clientInfo.OrderDate}");
            }
        }

        private void PrintDisplayAllClients(List<Client> clients)
        {
            Console.WriteLine($"код клиента\tНазвание организации\tАдрес\tФИО");
            foreach (var clientInfo in clients)
            {
                Console.WriteLine($"{clientInfo.Code}\t{clientInfo.OrganizationName}\t{clientInfo.Address}\t{clientInfo.ContactPerson}");
            }
        }
    }
}
