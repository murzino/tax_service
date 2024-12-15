## Как запустить проект
1. Клонируйте или скачайте репозиторий:
   git clone https://github.com/murzino/tax_service.git

2. Перейдите в директорию проекта:
   cd tax_service

3. Соберите Docker-образ:
    docker build -t tax_servise_progect .

4. Запустите контейнер:
    docker run -d -p 8000:8000 tax_servise_progect

5. Перейдите по адресу http://localhost:8000 в вашем браузере.

6. Нажмите кнопку 'Выберите файл' и выберете файл с исходными данными 'example_data.xlsx'.

7. Нажмите кнопку 'Загрузить' файл будет скачан.

