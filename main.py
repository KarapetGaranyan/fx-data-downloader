import downloader, data


# Скачивание и конвертация файлов
downloader.get_files()
downloader.files_converter()
data.get_fut_fin()

# Обработка и сохранение данных
data_EUR = data.get_data("EUR", 'OPTION TYPE: Monthly Options', 10000, data.get_close_prices()[0])
data_GBP = data.get_data("GBP", 'OPTION TYPE: Monthly Options', 1000, data.get_close_prices()[1])
data_AUD = data.get_data("AUD", 'OPTION TYPE: Monthly Options', 10000, data.get_close_prices()[2])
data_CAD = data.get_data("CAD", 'OPTION TYPE: Monthly Options', 10000000, data.get_close_prices()[3])
data_JPY = data.get_data("JPY", 'OPTION TYPE: Monthly Options', 1000000, data.get_close_prices()[4])
data_XAU = data.get_data("XAU", 'OPTION TYPE: American Options', 1, data.get_close_prices()[5])
data_XAG = data.get_data("XAG", 'OPTION TYPE: American Options', 100, data.get_close_prices()[6])

