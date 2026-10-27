# ms-store-download

Download packages directly from the Microsoft Store without a Windows machine.

The script uses the Windows Update SOAP API (the same protocol Windows itself uses) to resolve time-limited CDN download URLs for `.msix`, `.appx`, `.msixbundle`, and `.appxbundle` packages — no third-party services required.

## Requirements

Python 3.7+ with no external dependencies (stdlib only).

## Finding a Product ID

Open the Store page for the app in a browser. The product ID is the last segment of the URL:

```
https://www.microsoft.com/store/apps/9ndh0f2vhzx2
                                      ^^^^^^^^^^^^
```

## Usage

```
python ms_store_download.py <product_id> [options]
```

### Print download URLs

```sh
# All packages (all architectures)
python ms_store_download.py 9ndh0f2vhzx2

# Filter to x64 only
python ms_store_download.py 9ndh0f2vhzx2 --architecture x64

# Detect current machine architecture automatically
python ms_store_download.py 9ndh0f2vhzx2 --architecture auto
```

Output:
```
-> Fetching download link: [==============================] 100%
FileName:     MyApp_1.2.3.0_x64.msix
Architecture: x64
Url:          https://tlu.dl.delivery.mp.microsoft.com/...
Digest:       abc123...== (SHA256)
```

### Check the latest version quickly

```sh
python ms_store_download.py 9ndh0f2vhzx2 --catalog
```

`--catalog` asks the Store's Display Catalog API instead of Windows Update: one request, well under a second, over verified TLS. It lists each package's version, size and SHA256, but no download URL.

```
Package:      MyApp_1.2.3.0_x64__8wekyb3d8bbwe
Architecture: x64
Version:      1.2.3.0
Format:       msix
Size:         1327687105
Digest:       abc123...== (SHA256)
```

### Download packages

```sh
# Download to ~/Downloads (default)
python ms_store_download.py 9ndh0f2vhzx2 --download

# Download to a specific directory
python ms_store_download.py 9ndh0f2vhzx2 --download /tmp/myapp

# Download x64 package for the current machine
python ms_store_download.py 9ndh0f2vhzx2 --architecture auto --download
```

Already-downloaded files are skipped if their SHA256 digest matches. An interrupted download is resumed from its `.part` file on the next run.

### Download and extract

```sh
# Extract into the same directory as the downloaded files (each in its own sub-directory)
python ms_store_download.py 9ndh0f2vhzx2 --download /tmp/myapp --extract

# Extract into a separate directory (each in its own sub-directory)
python ms_store_download.py 9ndh0f2vhzx2 --download /tmp/myapp --extract /tmp/myapp-extracted

# Extract directly into the extract directory, without creating a sub-directory per package
python ms_store_download.py 9ndh0f2vhzx2 --download /tmp/myapp --extract /tmp/myapp-extracted --extract-flat
```

By default each package is extracted into its own sub-directory named after the package file stem (e.g. `/tmp/myapp-extracted/MyApp_1.2.3.0_x64/`). Use `--extract-flat` to extract files directly into the extract directory instead.

### Include dependency packages

By default only the main app package is returned. Pass `--all-packages` to also include bundled framework and dependency packages:

```sh
python ms_store_download.py 9ndh0f2vhzx2 --all-packages --download /tmp/myapp
```

## Options

| Option | Default | Description |
|---|---|---|
| `product_id` | *(required)* | Microsoft Store product ID |
| `--architecture` | *(all)* | Filter by CPU architecture: `x86`, `x64`, `arm`, `arm64`, `neutral`, `auto` |
| `--all-packages` | off | Include bundled dependency packages |
| `--locale` | `en-US` | Device locale reported to the WU service |
| `--os-sku-id` | `48` | Windows SKU ID reported to the WU service (48 = Windows 10 Pro) |
| `--os-version` | `10.0.16184.1001` | Windows version string reported to the WU service |
| `--download [DIR]` | `~/Downloads` | Download packages to DIR |
| `--extract [DIR]` | *(download dir)* | Extract downloaded packages to DIR (requires `--download`) |
| `--extract-flat` | off | Extract files directly into the extract DIR instead of creating a sub-directory per package |
| `--catalog` | off | Only list versions, sizes and digests from the Display Catalog API (fast, no download URL) |

## Use as a library

```python
import pathlib
import ms_store_download as msd

latest = msd.get_catalog_packages("9ndh0f2vhzx2", architecture="x64")   # fast version check
# The newest version of one package: the catalog, or Windows Update if it can't answer
newest = msd.get_latest_package("9ndh0f2vhzx2", "Publisher.App", architecture="x64")
pkg = msd.get_download_links("9ndh0f2vhzx2", architecture="x64", progress_callback=None)[0]
path = msd.download_package(pkg, pathlib.Path("/tmp/pkgs"), lambda done, total: ...)
msd.extract_package(path, pathlib.Path("/tmp/app"), lambda done, total: ...)
```

- Every function raises `ms_store_download.StoreError` on failure instead of exiting.
- `download_package` resumes a `.part` file left by an interrupted download, and keeps it if the progress callback raises (to cancel, for instance). A file whose SHA256 doesn't match is deleted.
- `extract_package` checks for free disk space first.
- `get_latest_package` returns `Version`, `Digest` and `Size`, plus a `Url` when the answer came from Windows Update (`Source == "soap"`). The catalog gives none: ask again with `source="soap"` to get one.
- `ms_store_download.TIMEOUT` (default 30 s) bounds every network operation.

## Notes

- Only **free** apps can be downloaded. The WU service restricts access to encrypted payloads for paid apps.
- Some apps use `.exe` installers rather than Appx/MSIX packages and cannot be downloaded with this tool.
- The Windows Update endpoints are reached without TLS certificate verification (their chain ends in a Microsoft root that common Linux trust stores lack), and download URLs are plain HTTP. Downloads are checked against their SHA256; for a digest obtained over verified TLS, take it from `--catalog` / `get_catalog_packages`.
- Download URLs are time-limited (typically a few hours). Re-run the script to get a fresh URL if a link has expired.
