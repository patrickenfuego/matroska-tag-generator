# matroska-tag-generator

![MediaInfo Example](https://i.imgur.com/s3ohFhp.png)

A CLI-based PowerShell script for automating Matroska tag files using the TMDB API. Cross-platform and compatible with PowerShell 5.1 or greater (linux/macOS users must install PowerShell 6/7).

## About

This script allows you to quickly generate XML tag files for media using the Matroska (mkv) container format.

If the [MKVToolNix](https://mkvtoolnix.download/) suite is installed, the metadata file will be automatically appended to the media file unless specified otherwise. To use this script, you must have a valid [TMDB API key](https://www.themoviedb.org/documentation/api), which can be obtained for free on their website.

Custom metadata tags can be viewed by software like [MediaInfo](https://mediaarea.net/en/MediaInfo), and can be used by metadata scrapers in software like [Plex](https://www.plex.tv/) or [Emby](https://emby.media/). They're also a nice way to quickly view media information without accessing the internet.

By default, the script pulls the following fields:

- TMDB ID
- IMDb ID
- Cast (top 5 billed)
- Writers with designations such as 'screenplay' or 'novel' (top 3 billed)
- Director(s) (top 2 billed)

## Parameters

> For more information on parameters, run: `Get-Help .\MatroskaTagGenerator.ps1 -Full`

| Name             | Description                                                                                                      |
| ---------------- | ---------------------------------------------------------------------------------------------------------------- |
| `Path`           | Path to the output file. Can be an existing mkv file or a new XML file (required). Positional parameter          |
| `APIKey`         | TMDB API key to use (optional)                                                                                   |
| `Title`          | Optional clean title to use for the API search                                                                   |
| `Properties`     | Additional metadata properties to include. If they do not exist/cannot be found, they will be skipped (optional) |
| `SkipProperties` | Skip some/all of the property fields generated automatically, such cast **Cast** (optional)                      |
| `NoMux`          | Switch flag to disable automatic muxing of tag files into the container (optional)                               |
| `AllowClobber`   | Switch flag to force overwriting an XML file if one already exists at that location (optional)                   |

## Usage

By default, the script uses the leaf of the output file name as the basis for the API search. For example, if a file named `C:\Movies\Ex Machina.mkv` is passed, 'Ex Machina' will be the search query. For files with extra details (i.e. `~/movies/Ex.Machina.2014.2160p.BluRay.mkv`), the script will attempt to extract & sanitize the base film title automatically. If this does not work, use the `-Title` parameter to pass a clean title instead.

If the output file passed via `-Path` is an mkv file and `mkvpropedit` is installed on the system, the script will automatically add the tag file to the container (the file must exist already); if `mkvpropedit` is not installed, a tag file will be created in the same directory as the mkv file. If an XML file path is passed via `-Path`, the script will generate the file and save it at that location instead.

When using your API key, you may pass it via the `-APIKey` parameter (useful when multiple keys exist), or simply assign that variable a default value by modifying line 85:

```PowerShell
[string]$APIKey = 'your_key_here'
```

When declaring your key this way, you don't need to specify it each time you run the script (unless you really want to).

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
