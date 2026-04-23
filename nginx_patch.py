import sys
content = open('/etc/nginx/sites-enabled/geolead').read()
if '/addin/' in content:
    print('Already patched')
    sys.exit(0)
patch = """    location /addin/ {
        proxy_pass http://127.0.0.1:5000;
        proxy_http_version 1.1;
        proxy_set_header Host $host;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        add_header Cache-Control 'no-cache, must-revalidate';
    }

    location / {"""
content = content.replace('    location / {', patch, 1)
open('/etc/nginx/sites-enabled/geolead', 'w').write(content)
print('PATCHED')
