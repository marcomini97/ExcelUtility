using ExcelUtilityConsole;
using ExcelUtilityLibrary;
using Microsoft.Extensions.Configuration;

ExcelSettingsLibrary.SetLicenceKey("Marco Comini");

var config = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory())
                                       .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                                       .Build();

var aziendaSettings = config.GetRequiredSection("Azienda").Get<AziendaSettings>();

UtilitySettings.ManageMenuApplication(aziendaSettings);