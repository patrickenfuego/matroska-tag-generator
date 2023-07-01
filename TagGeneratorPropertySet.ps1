$propertySets = [PSCustomObject]@{
    # movie = 'IMDbID', 'TMDbID', 'Production Companies', 'Spoken Languages', 'Tagline', 'Production Countries',
    #         'Revenue', 'Runtime', 'Overview'
    # credits = 'Writers', 'Writer', 'Directors', 'Director', 'Cast'
    
    'Production Companies' = @{ endpoint = 'movie'; property = 'production_companies'; subProperty = 'name'; type = 'array' }
    'Production Countries' = @{ endpoint = 'movie'; property = 'production_countries'; subProperty = 'name'; type = 'array' }
    'Release Date'         = @{ endpoint = 'movie'; property = 'release_date'; format = '{0:MM/dd/yyy}'; type = 'date' }
    'Revenue'              = @{ endpoint = 'movie'; property = 'revenue'; format = '{0:C}'; type = 'currency' }
    'Runtime'              = @{ endpoint = 'movie'; property = 'runtime' }
    'Spoken Languages'     = @{ endpoint = 'movie'; property = 'spoken_languages' }
}

[scriptblock]$formatCurrency = { 
    param ($Currency)

    return '{0:C}' -f $Currency
}

$propertySets | Add-Member -MemberType ScriptMethod -Name FormatCurrency -Value $formatCurrency
$propertySets | gm
$propertySets.FormatCurrency(1000000)


# Help organize which endpoint to use for custom properties
$propertySets = [PSCustomObject]@{
    # movie = 'IMDbID', 'TMDbID', 'Production Companies', 'Spoken Languages', 'Tagline', 'Production Countries',
    #         'Revenue', 'Runtime', 'Overview'
    # credits = 'Writers', 'Writer', 'Directors', 'Director', 'Cast'
    
    'Production Companies' = @{ endpoint = 'movie'; property = 'production_companies'; subProperty = 'name'; type = 'array' }
    'Production Countries' = @{ endpoint = 'movie'; property = 'production_countries'; subProperty = 'name'; type = 'array' }
    'Release Date'         = @{ endpoint = 'movie'; property = 'release_date'; format = '{0:MM/dd/yyy}'; type = 'date' }
    'Revenue'              = @{ endpoint = 'movie'; property = 'revenue'; format = '{0:C}'; type = 'currency' }
    'Runtime'              = @{ endpoint = 'movie'; property = 'runtime' }
    'Spoken Languages'     = @{ endpoint = 'movie'; property = 'spoken_languages'; subProperty = 'name'; type = 'array' }
    'Tagline'              = @{ endpoint = 'movie'; property = 'tagline' }
}
