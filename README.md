# Docket — Word ⇄ PDF converter website

A working iLovePDF-style converter: a FastAPI backend (LibreOffice + pdf2docx
do the real conversion work) and a plain HTML/JS frontend. Tested end-to-end:
docx→pdf and pdf→docx both verified to preserve content correctly.

```
webapp/
├── backend/
│   ├── main.py           # FastAPI app: /api/to-pdf, /api/to-word
│   ├── converter.py       # conversion logic (LibreOffice + pdf2docx)
│   ├── requirements.txt
│   └── Dockerfile
├── frontend/
│   └── index.html         # upload UI, drag & drop, calls the API
└── docker-compose.yml      # runs backend + frontend together
```

---

## 1. Get a Hostinger VPS (not shared hosting)

Shared/cPanel hosting won't work — you can't install LibreOffice there.
You need a **VPS plan** (Hostinger's cheapest VPS tier is enough to start).

1. Buy a Hostinger VPS plan, choose **Ubuntu 22.04** as the OS template.
2. Note the VPS's public IP address and root password (emailed to you /
   shown in hPanel).
3. Point your domain's **A record** at the VPS IP (in your domain's DNS
   settings — can be Hostinger's DNS or wherever your domain is registered).

## 2. Connect to your VPS

On Windows, use **PowerShell** or download **PuTTY**:

```
ssh root@YOUR_VPS_IP
```

Enter the password when prompted.

## 3. Install Docker on the VPS

Run these commands on the VPS (paste one at a time):

```bash
apt update && apt upgrade -y
curl -fsSL https://get.docker.com -o get-docker.sh
sh get-docker.sh
apt install docker-compose-plugin -y
docker --version
```

You should see a Docker version printed.

## 4. Upload the project files

From your **local Windows machine**, open PowerShell in the folder containing
the `webapp` folder and run (replace the IP):

```
scp -r webapp root@YOUR_VPS_IP:/root/
```

This copies the whole project to the VPS. (If `scp` isn't recognized, install
[WinSCP](https://winscp.net) instead and drag-and-drop the `webapp` folder to
`/root/` on the server through its GUI.)

## 5. Build and run

Back in your SSH session on the VPS:

```bash
cd /root/webapp
docker compose up -d --build
```

This will:
- Build the backend image (installs LibreOffice + Python deps inside it)
- Start the backend on port 8000
- Start an nginx server serving the frontend on port 80

Check it's running:

```bash
docker compose ps
curl http://localhost:8000/api/health
```

You should see `{"status":"ok"}`.

## 6. Point the frontend at your real domain

Edit `frontend/index.html` on the VPS (or before uploading) and change:

```js
const API_BASE = "http://localhost:8000";
```

to your actual domain, e.g.:

```js
const API_BASE = "https://api.yourdomain.com";
```

Then rebuild the frontend container:

```bash
docker compose up -d --build frontend
```

## 7. Add HTTPS (required — browsers block plain-http uploads on real domains)

Install Certbot and get a free SSL certificate:

```bash
apt install nginx certbot python3-certbot-nginx -y
```

The simplest path: put **Nginx directly on the host** (not in Docker) as a
reverse proxy in front of both containers, then let Certbot manage certs for
it. A minimal `/etc/nginx/sites-available/docket` config:

```nginx
server {
    listen 80;
    server_name yourdomain.com;

    location / {
        proxy_pass http://localhost:80;  # frontend container
    }
}

server {
    listen 80;
    server_name api.yourdomain.com;

    location / {
        proxy_pass http://localhost:8000; # backend container
        client_max_body_size 30M;
    }
}
```

```bash
ln -s /etc/nginx/sites-available/docket /etc/nginx/sites-enabled/
nginx -t && systemctl reload nginx
certbot --nginx -d yourdomain.com -d api.yourdomain.com
```

Certbot will auto-configure HTTPS and set up auto-renewal.

## 8. Restrict CORS (recommended before going live)

In `backend/main.py`, change:

```python
allow_origins=["*"],
```

to:

```python
allow_origins=["https://yourdomain.com"],
```

Rebuild: `docker compose up -d --build backend`

---

## Updating the site later

```bash
cd /root/webapp
# upload new files (scp) then:
docker compose up -d --build
```

## Notes on limits

- Current upload cap is 25MB per file (`MAX_FILE_SIZE_MB` in `main.py`) —
  raise it if you need bigger documents, and also raise `client_max_body_size`
  in the nginx config to match.
- Each conversion runs in an isolated temp folder that's deleted right after
  the file is sent back — nothing accumulates on disk.
- For heavier traffic later, put a queue (e.g. Redis + a worker) in front of
  conversions so uploads don't pile up on one process — not needed to launch.
