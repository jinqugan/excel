# Excel Setting

1. composer install
2. cp .env.example to .env
    - edit db setting
3. php artisan key:generate

# Excel Formatting

1. php artisan excel:dir
   - pre generate an empty folder if does not exist
2. php artisan format:multi
   - will modify all the excels file located at app\storage\excels\incomplete folder to      requested new excel format
   - formatted file will be generate at app\storage\excels\completed folder