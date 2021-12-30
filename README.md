# matroska-tag-generator

![MediaInfo Example](https://i.imgur.com/s3ohFhp.png)

A CLI-based PowerShell script for automating Matroska tag files using the TMDB API. Cross-platform and compatible with PowerShell 5.1 or greater (linux/macOS users must install PowerShell 6/7). 

Also included is a Windows `.bat` script that you can drag and drop files onto via the GUI.

## About

> NOTE: Using the `.bat` method for generating files only provides access to the `-Path` parameter, and may return incorrect results depending on the file name. Use the PowerShell script directly for full functionality

This script allows you to quickly generate XML tag files for media using the Matroska (mkv) container format.

If the [MKVToolNix](https://mkvtoolnix.download/) suite is installed, the metadata file will be automatically appended to the mkv file unless specified otherwise. If an XML file is passed, a tag file will be generated at that location.

To use this script, you must have a valid [TMDB API key](https://www.themoviedb.org/documentation/api), which can be obtained for free on their website.

Custom metadata tags can be viewed by software like [MediaInfo](https://mediaarea.net/en/MediaInfo), and can be used by metadata scrapers in software like [Plex](https://www.plex.tv/) or [Emby](https://emby.media/). They're also a nice way to quickly view media information without accessing the internet.

By default, the script pulls the following fields, but any/all of them can be omitted via parameters:

- TMDB ID
- IMDb ID
- Cast (top 5 billed)
- Writers with designations such as 'screenplay' or 'novel' (top 3 billed)
- Director(s) (top 2 billed)

## Parameters

> For more information on parameters, run: `Get-Help .\MatroskaTagGenerator.ps1 -Full`

| Name             | Description                                                                                                                     | Type       |
| ---------------- | ------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| `Path`           | Path to the output file. Can be an existing mkv file or a new XML file (required). This parameter is exposed to the `.bat` file | Positional |
| `APIKey`         | TMDB API key to use (optional as a passed parameter, but required in general)                                                   | Positional |
| `Title`          | Optional clean title to use for the API search                                                                                  | Positional |
| `Year`           | Optional year to use for the API search. Useful for differentiating titles with the same name (such as remakes)                 | Named      |
| `Properties`     | Additional metadata properties to include. If they do not exist/cannot be found, they will be skipped (optional)                | Named      |
| `SkipProperties` | Skip some/all of the property fields generated automatically, such **Cast** (optional)                                          | Named      |
| `NoMux`          | Switch flag to disable automatic muxing of tag files into the container (optional)                                              | Named      |
| `AllowClobber`   | Switch flag to force overwrite an XML file if one already exists                                                                | Named      |

## Usage

By default, the script uses the leaf of the output file name as the basis for the API search. For example, if a file named `C:\Movies\Ex Machina 2014.mkv` is passed, 'Ex Machina' and 2014 will be used in the search query. For files with extra details (i.e. `~/movies/Ex.Machina.2014.2160p.UHD.BluRay.mkv`), the script will attempt to extract & sanitize the base film title and year automatically. Title's that have a year in the name (Such as *Blade Runner 2049*) may not get parsed correctly, and the `-Year` parameter should be used.

If a year is not present (either via the file name or parameter), the script will grab the first object returned (usually the most relevant). If this does not work (or returns an incorrect match), use the `-Title` and/or `-Year` parameters to pass a clean title and release year instead.

When using your API key, you may pass it via the `-APIKey` parameter (useful when multiple keys exist), or simply assign that variable a default value by modifying line 85:

```PowerShell
[string]$APIKey = 'your_key_here',
```

## Examples

### Pass an mkv file to the script and append the tags automatically (if `mkvpropedit` is installed)

```PowerShell
PS > .\MatroskaTagGenerator.ps1 -Path 'C:\Movies\Ex.Machina.2014.UHD.2160p.HDR.bluray.mkv'
#Or, use the -Path parameter positionally
PS > .\MatroskaTagGenerator.ps1 '~/Movies/Ex.Machina.2014.UHD.2160p.HDR.bluray.mkv'
#Call the script from a different script or directory using the PowerShell & operator
PS > & ~/scripts/MatroskaTagGenerator.ps1 -Path '~/Movies/Ex.Machina.2014.UHD.2160p.HDR.bluray.mkv'
```

### Pass a clean title if results are not being returned

> NOTE: The script will attempt to extract & sanitize the title automatically

```PowerShell
PS > .\MatroskaTagGenerator.ps1 '~/Movies/Ex.Machina.2014.UHD.2160p.HDR.bluray.mkv' -Title 'Ex Machina'
```

### Pass an XML file to be created at the specified path

```PowerShell
PS > .\MatroskaTagGenerator.ps1 -Path 'C:\Movies\Ex Machina\Ex Machina.xml'
```

### Generate a tag file without Cast and IMDb ID metadata

```PowerShell
PS > .\MatroskaTagGenerator.ps1'C:\Movies\Ex Machina\Ex Machina.mkv' -SkipProperties cast, IMDbID
```

### Generate a tag file with additional metadata such as Overview and Budget

> NOTE: Some properties require deeper enumeration and may not get copied without modification

```PowerShell
PS > .\MatroskaTagGenerator.ps1'~/Movies/Ex Machina/Ex Machina.mkv' -Properties budget, overview
```

See the TMDB API documentation for a full list of properties.
