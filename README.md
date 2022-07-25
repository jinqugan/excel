# Marketplace

1. composer install
2. cp .env.example to .env
    - edit db setting, APP_URL
3. php artisan key:generate
4. php artisan db:seed
5. php artisan queue:work
   - edit mail setting (eg. mailtrap) in .env file to receive email notification

# Swagger Documentations

1. php artisan l5-swagger:generate
2. open documents with the following link ${APP_URL}/api/docs
   - eg. http://marketplace.test/api/docs