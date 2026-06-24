# Deployment Hardening Guide

This guide provides instructions for hardening the deployment of the SCD Dbase Sorter application.

## 1. Nginx Reverse Proxy with HTTPS

It is strongly recommended to run the Streamlit dashboard behind an Nginx reverse proxy. This provides SSL/TLS termination, better performance, and an additional security layer.

### Nginx Configuration Example
Create a file at `/etc/nginx/sites-available/scd-sorter`:

```nginx
server {
    listen 80;
    server_name yourdomain.com;
    return 301 https://$server_name$request_uri;
}

server {
    listen 443 ssl;
    server_name yourdomain.com;

    ssl_certificate /etc/letsencrypt/live/yourdomain.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/yourdomain.com/privkey.pem;

    # Strong SSL settings
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_prefer_server_ciphers on;
    ssl_ciphers 'ECDHE-ECDSA-AES128-GCM-SHA256:ECDHE-RSA-AES128-GCM-SHA256:ECDHE-ECDSA-AES256-GCM-SHA384:ECDHE-RSA-AES256-GCM-SHA384:ECDHE-ECDSA-CHACHA20-POLY1305:ECDHE-RSA-CHACHA20-POLY1305:DHE-RSA-AES128-GCM-SHA256:DHE-RSA-AES256-GCM-SHA384';

    location / {
        proxy_pass http://localhost:8501;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;

        # Streamlit-specific settings
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_read_timeout 86400;
    }
}
```

## 2. SSL/TLS Certificate (Let's Encrypt)

Use `certbot` to obtain and automatically renew certificates:

```bash
sudo apt update
sudo apt install certbot python3-certbot-nginx
sudo certbot --nginx -d yourdomain.com
```

## 3. Firewall Configuration (UFW)

Restrict incoming traffic to only necessary ports:

```bash
sudo ufw default deny incoming
sudo ufw default allow outgoing
sudo ufw allow ssh
sudo ufw allow 'Nginx Full'
sudo ufw enable
```

## 4. Intrusion Prevention (Fail2Ban)

Install and configure `fail2ban` to protect against brute-force attacks:

```bash
sudo apt install fail2ban
```

Create `/etc/fail2ban/jail.local`:

```ini
[sshd]
enabled = true
port = ssh
filter = sshd
logpath = /var/log/auth.log
maxretry = 3
bantime = 1h

[nginx-http-auth]
enabled = true
filter = nginx-http-auth
port = http,https
logpath = /var/log/nginx/error.log
```

## 5. Environment Variable Management

Do NOT hardcode secrets (like SMTP passwords or Master Keys) in the code. Use a `.env` file and load them using `python-dotenv`.

Example `.env` file:
```env
SMTP_PASSWORD=your_secure_password
MASTER_KEY_PATH=/path/to/secure/key
```

Ensure `.env` is added to `.gitignore`.

## 6. Directory Permissions

Ensure the `data/` directory has restricted permissions:

```bash
# Data directory should be accessible only by the application user
chmod 700 /path/to/SCD_Dbase_Sorter/data
chmod 600 /path/to/SCD_Dbase_Sorter/data/config/.master.key
```
