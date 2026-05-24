# Backups

Two state stores to back up:

1. **Workbook files** — what users actually work on. Lives in
   whatever `CASUAL_STORAGE` backend you chose.
2. **Admin config** — branding + webhooks + storage creds. Lives at
   `CASUAL_ADMIN_CONFIG_PATH` (default `/data/casual-admin.json`).

The Y.Doc room state in Redis is **transient**. Don't bother backing
it up — clients reconnect, the workbook re-seeds from the storage
backend, and live edits resume from there. Redis is acceleration,
not storage.

## Per-backend recipes

### `CASUAL_STORAGE=local`

Workbooks at `$CASUAL_LOCAL_PATH` as `<fileId>.xlsx` +
`<fileId>.meta.json` pairs.

```sh
# Nightly snapshot to a sibling host.
rsync -av --delete /var/lib/casual-sheets/data/ \
  backup@dr.acme.example:/srv/backups/casual-sheets/$(date +%F)/
```

Restore = `rsync` the other way. Filesystem is the format —
no extraction needed.

### `CASUAL_STORAGE=s3`

Use the cloud provider's built-in versioning + lifecycle. AWS S3:

```sh
aws s3api put-bucket-versioning \
  --bucket casual-sheets \
  --versioning-configuration Status=Enabled

aws s3api put-bucket-lifecycle-configuration \
  --bucket casual-sheets \
  --lifecycle-configuration file://lifecycle.json
```

`lifecycle.json` keeping the last 30 days of non-current versions:

```json
{
  "Rules": [{
    "ID": "expire-old-versions",
    "Status": "Enabled",
    "Filter": {},
    "NoncurrentVersionExpiration": { "NoncurrentDays": 30 }
  }]
}
```

MinIO + R2 + B2 have equivalent UI / API toggles. R2 ships
versioning + lifecycle as opt-ins via the dashboard.

Restore = `aws s3api list-object-versions ... | aws s3api restore-object ...`.

### `CASUAL_STORAGE=postgres`

```sh
pg_dump --format=custom \
  $CASUAL_PG_URL \
  --table=casual_workbooks \
  > /backup/casual-sheets-$(date +%F).pgcustom
```

For incremental: WAL archiving via `archive_command` is the
standard recipe. See the Postgres docs.

Restore = `pg_restore`:

```sh
pg_restore --dbname=$CASUAL_PG_URL \
  --table=casual_workbooks \
  /backup/casual-sheets-2026-06-01.pgcustom
```

## Admin config

Whichever backend you use, the admin config JSON is on the
**filesystem inside the container** (default `/data/casual-admin.json`).
Bind-mount `/data` so it lands on a host volume + back the volume up
with rsync.

The file is mode `0600` and contains secrets (S3 secret key, OIDC
client secret, webhook signing secrets). Treat the backups the
same way — encrypt at rest, restrict access.

## Restore drill — do it once

Run through the restore at least once before you need it:

1. Spin up a clean container in a side network with the backup
   volume mounted.
2. `docker compose up -d` against the restored data.
3. Open the admin panel → verify config (branding, backend
   selection, room limits).
4. Open a workbook by id from the storage backend — confirm
   bytes match.
5. Tear down the side network.

The Y.Doc state Redis doesn't restore — but the workbook bytes
do, and clients reconnect into fresh rooms seeded from those
bytes. No data loss; brief session-state loss for in-flight
edits at the moment of restore.

## What's NOT in any backup

| Item | Where it lives | Recovery |
|---|---|---|
| `CASUAL_JWT_SECRET` | Your deployment manifest / secret manager. | Keep a copy. Losing it invalidates every issued token. |
| `CASUAL_ADMIN_PASSWORD` | Same. | Generate a new one + re-login. |
| Redis WAL | The Redis volume. | Lost on Redis volume loss; clients reconnect + rebuild. |
| In-flight WebSocket sessions | RAM. | Clients reconnect. |
