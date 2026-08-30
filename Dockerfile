# Use official PHP 8.2 Apache base image
FROM php:8.2-apache

# Install libzip-dev and zip tools, then compile the zip extension
# (Note: simplexml, xml, and dom are already pre-compiled in php:8.2-apache)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libzip-dev \
    zip \
    unzip \
    && docker-php-ext-install -j$(nproc) zip \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Enable Apache mod_rewrite
RUN a2enmod rewrite

# Configure PHP limits for processing large batches
RUN { \
    echo 'upload_max_filesize = 128M'; \
    echo 'post_max_size = 128M'; \
    echo 'max_execution_time = 600'; \
    echo 'memory_limit = 512M'; \
} > /usr/local/etc/php/conf.d/word-texter.ini

# Set working directory
WORKDIR /var/www/html

# Copy application files into container
COPY . /var/www/html/

# Set proper permissions for Apache user
RUN mkdir -p /var/www/html/secure_data \
    && chown -R www-data:www-data /var/www/html \
    && chmod -R 755 /var/www/html \
    && chmod -R 775 /var/www/html/secure_data

# Expose HTTP port
EXPOSE 80

# Start Apache web server
CMD ["apache2-foreground"]
