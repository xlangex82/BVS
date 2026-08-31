# BVS
Self-Hosted **BewerbungsVerwaltungsSystem** basierend auf Docker

B-V-S is a self-hosted Docker application for managing job applications, CVs
and cover letters. It is designed primarily for private, self-hosted use.

**Current release: v8.5.5**

## Features

- Application management with configurable statuses
- Dashboard, overview, Kanban board and calendar
- Follow-up dates and in-app reminders
- Active/completed process filtering and Excel export
- Master CV, universal CVs and application-specific CV snapshots
- Configurable CV templates, browser previews and PDF generation
- Master and application-specific cover letters
- Cover-letter templates, signatures and certification-logo footer
- Dynamic A4 cover-page designer
- Application history, branding/personalization and backup

## Stack

Python / Flask, Jinja, SQLAlchemy, MySQL 8.4, ReportLab, vanilla CSS/JavaScript
and Docker Compose.

## Requirements

Docker with the Docker Compose plugin.

## Quick start

```bash
git clone https://github.com/xlangex82/BVS.git
cd b-v-s
docker compose up -d --build
```

For upgrades:

```bash
docker compose down
docker compose up -d --build
```

**Never delete `./mysql` and do not use `docker compose down -v` if you want to
retain your database.**

## Persistent data

Runtime data is intentionally outside Git:

```text
./mysql
./data/attachments
```

The first directory contains the MySQL data. Attachments include uploaded
documents, branding, CV media, signatures and cover-letter media.

## Backup

Use the backup function in the B-V-S web interface and keep generated backups
outside the Git repository.

## Security

B-V-S intentionally has **no built-in authentication**. It is intended for a
trusted private environment. For Internet-facing use, place it behind an
authenticated reverse proxy with HTTPS/TLS and appropriate firewall rules.

Applications and CVs contain personal data. Never commit databases,
attachments, backups, signatures, photos, generated PDFs, exports, credentials
or a real `.env` file.

See [SECURITY.md](SECURITY.md).

## CV model

B-V-S supports three CV levels:

1. **Master CV** – central source data.
2. **Universal CV** – reusable snapshot independent of an application.
3. **Application CV** – independent snapshot for a specific application.

## Contributing

Contributions are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

B-V-S is released under the [MIT License](LICENSE).

Copyright © 2026 Lange-IT.com / Peter Lange.

## Author

Peter Lange — Lange-IT.com
