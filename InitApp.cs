using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OfficeOpenXml;
using Renga;
using RengaFacade;

namespace RengaTemplate
{
    public class InitApp : IPlugin
    {
        ActionEventSource action;

        [DllImport("kernel32")]
        static extern bool AllocConsole();

        public bool Initialize(string pluginFolder)
        {
            // запуск консоли ренги
            AllocConsole();

            #region Инициализация плагина
            var rengaApp = new Renga.Application();
            var rengaUI = rengaApp.UI;
            var panel = rengaUI.CreateUIPanelExtension();

            var button = rengaUI.CreateAction();
            button.ToolTip = "Обработка копий свойств";

            var icon = rengaUI.CreateImage();
            icon.LoadFromFile($@"{pluginFolder}\Resources\logo.png");
            button.Icon = icon;
            #endregion

            action = new ActionEventSource(button);
            action.Triggered += (sender, args) =>
            {
                try
                {
                    Console.ForegroundColor = ConsoleColor.White;

                    #region описание

                    Console.WriteLine("Плагин \"Property Copies Resolver\" создает свойства и переносит в них нестандартные значения из существующих свойств.");
                    Console.WriteLine("Данные для сопоставления свойств читаются из таблицы Excel. Она должна иметь следующий вид:");
                    Console.WriteLine("");
                    Console.WriteLine("+---+------------+------------+");
                    Console.WriteLine("|   | A          | B          |");
                    Console.WriteLine("+---+------------+------------+");
                    Console.WriteLine("| 1 | Свойство 1 | Свойство 2 |");
                    Console.WriteLine("+---+------------+------------+");
                    Console.WriteLine("| 2 | Свойство 3 |            |");
                    Console.WriteLine("+---+------------+------------+");
                    Console.WriteLine("");
                    Console.WriteLine("По данным из строки 1:");
                    Console.WriteLine("Будет создано свойство с именем \"Свойство 1\". Типы объектов будут скопированы из существующих свойств с именами \"Свойство 1\" и \"Свойство 2\".");
                    Console.WriteLine("Значением созданного свойства станет первое найденное нестандартное значение из свойств с именами \"Свойство 2\" (в пределах объекта).");
                    Console.WriteLine("");
                    Console.WriteLine("По данным из строки 2:");
                    Console.WriteLine("Будет создано свойство с именем \"Свойство 3\". Типы объектов будут скопированы из существующих свойств с именем \"Свойство 3\".");
                    Console.WriteLine("Значением созданного свойства станет первое найденное нестандартное значение из свойств с именами \"Свойство 3\" (в пределах объекта).");
                    Console.WriteLine("");
                    Console.WriteLine("Дополнительно:");
                    Console.WriteLine(" - Все свойства, неуказанные в столбце А, будут удалены;");
                    Console.WriteLine(" - Все копии свойств из столбца А будут удалены;");
                    Console.WriteLine(" - Значения в столбце А должны быть уникальными;");
                    Console.WriteLine(" - \"Выражение для свойства\" игнорируется;");
                    Console.WriteLine(" - Флаг \"Экспортировать значение свойства в CSV\" игнорируется.");
                    Console.WriteLine("");

                    MessageBox.Show("Требуется действие в консоли ренги", "Property Copies Resolver", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    while (true)
                    {
                        WriteBlue("Продолжить? [y/n] ");
                        var value = Console.ReadLine();
                        if (value == "y") break;
                        if (value == "n") return;
                        WriteLineYellow("Допустимые значения: \"y\" - да, \"n\" - нет");
                    }

                    #endregion

                    #region получение и чтение таблицы

                    var rulesPath = string.Empty;
                    while (true)
                    {
                        WriteBlue("Путь до таблицы: ");
                        rulesPath = Console.ReadLine();

                        if (!File.Exists(rulesPath))
                        {
                            WriteLineYellow("Файл по указанному пути не существует");
                            continue;
                        }
                        if (Path.GetExtension(rulesPath) != ".xlsx")
                        {
                            WriteLineYellow("Файл должен иметь расширение .xlsx");
                            continue;
                        }
                        break;
                    }

                    Console.WriteLine("Чтение таблицы...");
                    var rules = new List<PropertyRule>();
                    using (var stream = new FileStream(rulesPath, FileMode.Open, FileAccess.Read))
                    {
                        var existingFile = new FileInfo(rulesPath);
                        using (var package = new ExcelPackage(existingFile))
                        {
                            var worksheet = package.Workbook.Worksheets[1];

                            for (int rowIndex = 1; rowIndex <= worksheet.Dimension.Rows; rowIndex++)
                            {
                                rules.Add(new PropertyRule()
                                {
                                    PropertyName = worksheet.Cells[rowIndex, 1].GetValue<string>(),
                                    GetValueFromName = worksheet.Cells[rowIndex, 2].GetValue<string>()
                                });
                            }
                        }
                    }

                    var tablePropNameCopies = new List<string>();
                    foreach (var rule in rules)
                    {
                        if (string.IsNullOrEmpty(rule.PropertyName)) continue;

                        var copies = rules.Where(x => x.PropertyName == rule.PropertyName);
                        if (copies.Count() > 1)
                        {
                            tablePropNameCopies.Add(rule.PropertyName);
                        }

                        if (string.IsNullOrEmpty(rule.GetValueFromName) &&
                           !string.IsNullOrEmpty(rule.PropertyName))
                        {
                            rule.GetValueFromName = rule.PropertyName;
                        }
                    }
                    if (tablePropNameCopies.Any())
                    {
                        WriteLineRed("Таблица составлена некорректно. Следующие свойства в столбце А встречаются более одного раза:");
                        foreach (var name in tablePropNameCopies)
                        {
                            WriteLineRed($" - \"{name}\"");
                        }
                        return;
                    }

                    #endregion

                    var project = rengaApp.Project;
                    var facade = new RengaFacade.RengaFacade(project);

                    #region чтение свойств проекта и сопоставление с таблицей

                    Console.WriteLine("Чтение определений свойств...");
                    var AskToContinue = false;
                    var rulesUpd = new List<PropertyRule>();
                    var props = facade.GetProperties();
                    foreach (var rule in rules)
                    {
                        if (string.IsNullOrEmpty(rule.PropertyName) &&
                            string.IsNullOrEmpty(rule.GetValueFromName)) continue;

                        var propA = props.FirstOrDefault(x => x.Name == rule.PropertyName);
                        var propB = props.FirstOrDefault(x => x.Name == rule.GetValueFromName);

                        if (propA != null && propB != null)
                        {
                            rulesUpd.Add(rule);
                        }

                        if (propA == null && propB != null)
                        {
                            WriteLineYellow($"В проекте нет свойства \"{rule.PropertyName}\", поэтому оно будет создано. Значения будут взяты из копий свойства \"{rule.GetValueFromName}\".");
                            rulesUpd.Add(new PropertyRule()
                            {
                                PropertyName = rule.PropertyName,
                                GetValueFromName = rule.GetValueFromName
                            });
                            AskToContinue = true;
                        }

                        if (propA != null && propB == null)
                        {
                            WriteLineYellow($"В проекте нет свойства \"{rule.GetValueFromName}\". Значения будут взяты из копий свойства \"{rule.PropertyName}\".");
                            rulesUpd.Add(new PropertyRule()
                            {
                                PropertyName = rule.PropertyName,
                                GetValueFromName = rule.PropertyName
                            });
                            AskToContinue = true;
                        }

                        if (propA == null && propB == null)
                        {
                            WriteLineYellow($"В проекте нет свойств  \"{rule.PropertyName}\" и \"{rule.GetValueFromName}\".");
                            AskToContinue = true;
                        }
                    }
                    rules = rulesUpd;

                    #endregion

                    #region сбор данных для создания новых свойств

                    var propNamesFromTable =
                        rules.Where(x => x.PropertyName != null).Select(x => x.PropertyName)
                        .Concat(
                        rules.Where(x => x.GetValueFromName != null).Select(x => x.GetValueFromName))
                        .Distinct()
                        .OrderBy(x => x);
                    
                    var propsFromTable = facade.GetProperties().Where(x => propNamesFromTable.Contains(x.Name));
                    foreach (var rule in rules)
                    {
                        IEnumerable<RProperty> propsWithValues;

                        // значение нужно забирать из копий
                        if (rule.PropertyName == rule.GetValueFromName)
                        {
                            propsWithValues = propsFromTable.Where(x => x.Name == rule.PropertyName);
                            rule.GetValueFrom = propsWithValues;
                            rule.PropertyObjectTypes = propsWithValues
                                .SelectMany(x => x.ObjectTypes)
                                .Distinct();
                        }

                        // значение нужно забирать из других свойств
                        else
                        {
                            propsWithValues = propsFromTable.Where(x => x.Name == rule.GetValueFromName);
                            rule.GetValueFrom = propsWithValues;
                            rule.PropertyObjectTypes = propsWithValues
                                .SelectMany(x => x.ObjectTypes)
                                .Concat(propsFromTable
                                    .Where(x => x.Name == rule.PropertyName)
                                    .SelectMany(x => x.ObjectTypes))
                                .Distinct();
                        }

                        var pType = rule.GetValueFrom.First().PropertyType;
                        foreach (var propWithValue in propsWithValues)
                        {
                            if (pType != propWithValue.PropertyType)
                            {
                                WriteLineYellow($"Копии свойства \"{rule.GetValueFromName}\" имеют различные типы данных. Новому свойству будет присвоен тип данных - \"Строка\".");
                                pType = PropertyType.PropertyType_String;
                                AskToContinue = true;
                            }
                        }
                        rule.PropertyType = pType;
                    }

                    if (AskToContinue)
                    {
                        while (true)
                        {
                            WriteBlue("Продолжить? [y/n] ");
                            var value = Console.ReadLine();
                            if (value == "y") break;
                            if (value == "n") return;
                            WriteLineYellow("Допустимые значения: \"y\" - да, \"n\" - нет");
                        }
                    }

                    #endregion

                    #region создание, заполнение, удаление свойств

                    Console.WriteLine("Создание новых свойств...");
                    foreach (var rule in rules)
                    {
                        rule.Property = facade.CreateProperty(rule.PropertyName, rule.PropertyType);
                        rule.Property.AddToObjectType(rule.PropertyObjectTypes.Select(x => x.Id));
                    }

                    var objs = facade.GetObjects();
                    var counter = 0;
                    Console.WriteLine("Обработка объектов...");
                    Console.WriteLine($"Количество объектов - {objs.Count()} шт.");
                    var operation = project.CreateOperation();
                    operation.Start();

                    foreach (var obj in objs)
                    {
                        foreach (var rule in rules)
                        {
                            var objProps = obj.Properties;
                            if (objProps.FirstOrDefault(x => x.Id == rule.Property.Id) is RProperty prop)
                            {
                                if (objProps.FirstOrDefault(x => x.Name == rule.GetValueFromName && !x.HasDefaultValue) is RProperty propWithValue)
                                {
                                    try
                                    {
                                        prop.Value = propWithValue.Value;
                                    }
                                    catch (Exception ex)
                                    {
                                        WriteLineRed(ex.Message);
                                    }
                                }
                            }
                        }
                        counter++;
                        Console.WriteLine($"[{DateTime.Now:T}] {Math.Round((double)100 * counter / objs.Count(), 2, MidpointRounding.ToEven)} %");
                    }

                    operation.Apply();

                    Console.WriteLine("Удаление лишних свойств...");
                    var requiredPropIds = rules.Select(x => x.Property.Id);
                    var deletePropIds = new List<Guid>();
                    foreach (var id in facade.GetProperties().Select(x => x.Id).ToArray())
                    {
                        if (!requiredPropIds.Contains(id))
                        {
                            deletePropIds.Add(id);
                        }
                    }

                    foreach (var id in deletePropIds)
                    {
                        try
                        {
                            facade.DeleteProperty(id);
                        }
                        catch (Exception ex)
                        {
                            WriteLineRed($"{ex.Message} | Свойство - {id}.");
                        }
                    }

                    WriteLineGreen("Готово!");

                    #endregion
                }
                catch (Exception ex)
                {
                    WriteLineRed(ex.Message);
                }
            };

            panel.AddToolButton(button);
            rengaUI.AddExtensionToPrimaryPanel(panel);

            return true;
        }

        public void Stop()
        {
            action.Dispose();
        }

        private void WriteLineGreen(string message)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White;
        }

        private void WriteLineYellow(string message)
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White;
        }

        private void WriteLineRed(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White;
        }

        private void WriteBlue(string message)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.Write(message);
            Console.ForegroundColor = ConsoleColor.White;
        }

        private void WriteLineBlue(string message)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White;
        }
    }
}
