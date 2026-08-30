# Use official PHP 8.2 Apache image
FROM php:8.2-apache

# Install libzip-dev and zip tools, then install and enable the zip extension.
# (Note: simplexml, dom, and xml are already built into php:8.2-apache by default)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libzip-dev \
    zip \
    unzip \
    && docker-php-ext-install -j$(nproc) zip \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Enable Apache mod_rewrite
RUN a2enmod rewrite

# Configure PHP runtime & file upload directives
RUN { \
    echo 'upload_max_filesize = 128M'; \
    echo 'post_max_size = 128M'; \
    echo 'max_execution_time = 600'; \
    echo 'memory_limit = 512M'; \
} > /usr/local/etc/php/conf.d/word-texter.ini

# Set working directory
WORKDIR /var/www/html

# Copy application files into the container
COPY . /var/www/html/

# Create and configure secure storage directory permissions
RUN mkdir -p /var/www/html/secure_data \
    && chown -R www-data:www-data /var/www/html \
    && chmod -R 755 /var/www/html \
    && chmod -R 775 /var/www/html/secure_data

# Expose standard web port
EXPOSE 80

# Start Apache in foreground
CMD ["apache2-foreground"]
