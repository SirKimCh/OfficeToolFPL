FROM nginx
COPY /assets/. /usr/share/nginx/html/assets/.
COPY nginx.conf /etc/nginx/conf.d/default.conf