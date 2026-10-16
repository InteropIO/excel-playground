FROM 389653476181.dkr.ecr.eu-central-1.amazonaws.com/utils:nginx-unprivileged

COPY  --chown=nginx:nginx /docs /usr/share/nginx/html

EXPOSE 8080