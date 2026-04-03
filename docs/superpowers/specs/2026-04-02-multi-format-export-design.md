# Multi-Format CRDB Export

## Summary

Replace the single-format CRDB export (user picks GPKG or CSV) with an automatic multi-format export that writes all five formats to a date-stamped GIS folder.

## Output Path

Walk up from the CRDB file's directory to find an `ASBUILT` folder, then create:

```
ASBUILT/0_GIS/WEEKLY_UPDATE/YYYYMMDD/
```

If `ASBUILT` is not found within 6 ancestor levels, fall back to `<crdb_dir>/0_GIS/WEEKLY_UPDATE/YYYYMMDD/`.

## Files Produced

All files share the CRDB's stem name (e.g., `MAVERICK_PW_PRELIMINARY__CONSTRUCTION_STAKING`):

| File | Format | Writer |
|------|--------|--------|
| `<stem>.gpkg` | GeoPackage | Existing `_write_gpkg` |
| `<stem>.csv` | CSV (RAW_POINTS) | Existing `_write_crdb_csv` |
| `<stem>.shp` + `.shx` + `.dbf` + `.prj` | Shapefile | New `_write_crdb_shp` |
| `<stem>.xml` | LandXML 1.2 | New `_write_crdb_landxml` |
| `<stem>.kmz` | KMZ | New `_write_crdb_kmz` |

## New Writers

### Shapefile (`_write_crdb_shp`)

Single combined file with all points. Schema:

- `point_name` (C:50), `N` (N:20.10), `E` (N:20.10), `Z` (N:20.10), `code` (C:30)
- `ATT_1` through `ATT_28` (C:80)
- `h_prec` (N:12.4), `v_prec` (N:12.4), `pdop` (N:8.2), `sats` (N:4.0)
- `method` (C:20), `media` (C:100), `src_jxl` (C:100)

Written using pure Python (struct-based .shp/.shx/.dbf/.prj) — no external dependencies. Geometry is WGS84 Point from JXL geodetic; null geometry for unmatched points.

### LandXML 1.2 (`_write_crdb_landxml`)

```xml
<?xml version="1.0" encoding="UTF-8"?>
<LandXML xmlns="http://www.landxml.org/schema/LandXML-1.2" version="1.2"
         date="YYYY-MM-DD" time="HH:MM:SS">
  <Project name="<stem>"/>
  <CgPoints>
    <CgPoint name="<point_name>" code="<code>">
      lat lon elev
    </CgPoint>
    ...
  </CgPoints>
</LandXML>
```

Coordinates: WGS84 lat/lon/height from JXL. Points without geodetic match use local N/E/Z. Written via `xml.etree.ElementTree`.

### KMZ (`_write_crdb_kmz`)

Zipped KML with one `<Folder>` per feature code, each containing `<Placemark>` elements. Extended data includes attributes, GNSS quality, and media filename. Written via `xml.etree.ElementTree` + `zipfile`.

## UI Changes

In `_show_crdb_export_dialog`:

1. Remove format radio buttons (`fmt_var`, `fmt_frame`)
2. Remove output file entry/browse (`out_var`, `out_entry`, `_browse_out`, `_on_format_change`)
3. Add read-only label showing the computed output folder path
4. Keep: DWG geometry checkbox, issues report checkbox, client schema selector, Browse JXL button

## Export Flow (`_do_export`)

1. Compute output directory and create it
2. Write all five formats sequentially (GPKG → CSV → SHP → LandXML → KMZ)
3. Each writer is independent — if one fails, log the error and continue with the rest
4. Show summary message listing all files written and any failures
5. Register in CRDB watchlist as before (using GPKG path)

## Watchlist / Headless Re-export

The watchlist stores the output directory base. Headless re-export regenerates all five formats into a new date-stamped folder.
