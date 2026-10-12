# WiseOffice static deployment

WiseOffice is served by the existing `openresty` container at
`https://trainwise.efrobot.com/wiseoffice/editor`. It mounts
`/home/app/wiseoffice` from the host; no separate editor container is needed.

From the `document` project root, build and prepare the subpath assets:

```sh
pnpm run build
python3 deployment/prepare-wiseoffice.py dist /tmp/wiseoffice-release
```

The preparation step scopes application, converter and service worker URLs to
`/wiseoffice/`, leaving the vendor engine unchanged. Validate with a real browser
under that prefix, including CSV import and document saving.

Upload into a staging directory and back up the live site and gateway config
before publishing. Replace files with temporary files and rename each into place;
publish HTML last. Retain previous hashed assets for already open tabs. Do not
rename the mounted site directory: Docker keeps the original directory mount.
Write configuration updates into the existing mounted config file, then run
`docker exec openresty nginx -t` and `docker exec openresty nginx -s reload`.

The deployment on 2026-10-12 has a backup at
`/home/app/wiseoffice-backups/20261012-1751/` and a verified file manifest at
`/home/app/wiseoffice-staging/20261012-1751.sha256`.

To roll back, copy the backup site's files through temporary files and rename
them into the live directory, with HTML last. Restore `default.conf` into
`/home/app/trainwise-openresty/conf.d/default.conf` in place, test and reload
OpenResty. Close editor tabs before reopening after a vendor version rollback.

The standalone engine does not bundle spellcheck dictionaries. Some large PPT
and ODP files can still fail conversion; adding a supported extension does not
guarantee conversion of every document.
