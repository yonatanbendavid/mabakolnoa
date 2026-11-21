# "Ma BaKolnoa" (what's at the movies): downloads the Lev, yes Planet, Cinema City and Movieland schedules, seat plans and trailer links,
# writes them into lev.html (one self-contained page), then opens it.
# Any part that fails to download keeps its previous data.
param([switch]$NoOpen)

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

$here       = Split-Path -Parent $MyInvocation.MyCommand.Path
$api        = 'https://ticket.lev.co.il/api'
$page       = Join-Path $here 'lev.html'
# On GitHub (the phone page's repository) the same page is called index.html
if (-not (Test-Path -LiteralPath $page) -and (Test-Path -LiteralPath (Join-Path $here 'index.html'))) { $page = Join-Path $here 'index.html' }
$utf8       = New-Object System.Text.UTF8Encoding $false
$stamp      = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

# A web client that gives up on a slow site after 20 seconds (the default waits 100 seconds per request)
# and accepts compressed replies, so one slow site can't hold up the whole update.
if (-not ('TimedWebClient' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Net;
public class TimedWebClient : WebClient {
    public int TimeoutMs = 20000;
    protected override WebRequest GetWebRequest(Uri address) {
        WebRequest r = base.GetWebRequest(address);
        if (r != null) {
            r.Timeout = TimeoutMs;
            HttpWebRequest h = r as HttpWebRequest;
            if (h != null) { h.ReadWriteTimeout = TimeoutMs; h.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate; }
        }
        return r;
    }
}
"@
}
# Progress lines for the window shown when the PC shortcut runs the update
function Step([string]$msg) { Write-Host ((Get-Date).ToString('HH:mm:ss') + '  ' + $msg) }
# Trailers found on earlier runs, read back from the page, so film pages are only fetched for new films
function Old-Trailers([string]$html, [string]$id) {
    $m = @{}
    $s = $html.IndexOf("<script id=`"$id`">")
    if ($s -ge 0) {
        $e = $html.IndexOf('</script>', $s)
        foreach ($x in [regex]::Matches($html.Substring($s, $e - $s), '"([^"]+)":\{"t":"(https://[^"]+)"')) { $m[$x.Groups[1].Value] = $x.Groups[2].Value }
    }
    return $m
}

function New-Client {
    $wc = New-Object TimedWebClient
    $wc.Encoding = [System.Text.Encoding]::UTF8
    $wc.Headers.Add('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)')
    $wc.Headers.Add('Accept', 'application/json')
    return $wc
}
function Get-Text([string]$url) { return [System.Text.Encoding]::UTF8.GetString((New-Client).DownloadData($url)) }
function Clean([string]$s) { return ($s -replace "[\\'`r`n<>]", ' ') }
# JSON never has "<" outside strings, and < is the same character inside them,
# so this keeps "</script>" in a film description from breaking the page.
$LT = '\' + 'u003c'   # the JSON escape for "<", built in two parts so editors keep it as text
function Safe([string]$json) { return $json.Replace('<', $LT) }
function Set-Block([string]$html, [string]$id, [string]$js) {
    $open  = "<script id=`"$id`">"
    $start = $html.IndexOf($open)
    if ($start -lt 0) { throw "lev.html is missing its $id block." }
    $from = $start + $open.Length
    $end  = $html.IndexOf('</script>', $from)
    if ($end -lt 0) { throw "lev.html has an unclosed $id block." }
    return $html.Substring(0, $from) + $js + $html.Substring($end)
}

$html = [System.IO.File]::ReadAllText($page, [System.Text.Encoding]::UTF8)
$scheduleError = $null; $hallsError = $null; $filmsError = $null; $json = $null

Step 'Lev schedule...'
# 1. Schedule for every Lev cinema (the page opens on Tel Aviv)
try {
    $json = Get-Text "$api/presentations?includeSynopsis=0"
    if (-not $json.TrimStart().StartsWith('{') -or $json -notmatch '"presentations"') {
        throw 'The cinema site returned something other than a schedule.'
    }
    $html = Set-Block $html 'lev-schedule' ("window.LEV_FETCHED = '$stamp'; window.LEV_DATA = " + (Safe $json) + ";")
}
catch { $scheduleError = Clean $_.Exception.Message; $json = $null }

Step 'Lev seat maps...'
# 2. Seat plans, one per hall in the schedule (decorative images removed)
if ($json) {
    try {
        $pairs = @{}
        foreach ($m in [regex]::Matches($json, '"venueId":\s*(\d+),[^{}]*?"seatplanId":\s*(\d+),[^{}]*?"venueLocationId":\s*(\d+)')) {
            # Location 1 is Lev's streaming service, which has no halls
            if ($m.Groups[3].Value -ne '1') { $pairs[$m.Groups[1].Value] = $m.Groups[2].Value }
        }
        if ($pairs.Count -eq 0) { throw 'No halls found in the schedule.' }
        $parts = @(); $missing = @()
        foreach ($venue in ($pairs.Keys | Sort-Object { [int]$_ })) {
            try {
                $wc = New-Client
                $wc.Headers.Add('Content-Type', 'application/json')
                $plan = $wc.UploadString("$api/seats/seatplanV2?venueId=$venue&seatplanId=$($pairs[$venue])", 'POST', '{}')
                if (-not $plan.TrimStart().StartsWith('{')) { throw 'not readable' }
                $plan = [regex]::Replace($plan, '"i":"data:[^"]*"', '"i":null')
                $parts += "`"$venue`": " + (Safe $plan)
            }
            catch { $missing += $venue }
        }
        if ($parts.Count -eq 0) { throw 'No seat plans could be downloaded.' }
        $html = Set-Block $html 'lev-halls' ("window.LEV_HALLS = {" + ($parts -join ', ') + "};")
        if ($missing.Count) { $hallsError = "no seat map for hall " + ($missing -join ', ') }
    }
    catch { $hallsError = Clean $_.Exception.Message }
}

Step 'Lev trailers...'
# 3. Film details, for trailer links (descriptions removed)
try {
    $films = Get-Text "$api/features"
    if (-not $films.TrimStart().StartsWith('[')) { throw 'Film details were not readable.' }
    $films = [regex]::Replace($films, '"(synopsis|additionalSynopsis)":"(?:[^"\\]|\\.)*"', '"$1":""')
    $html = Set-Block $html 'lev-films' ("window.LEV_FILMS = " + (Safe $films) + ";")
}
catch { $filmsError = Clean $_.Exception.Message }

Step 'yes Planet...'
# 3b. Planet cinemas: the next 8 days of screenings at every branch, stored compactly
$planetError = $null
try {
    $pb = 'https://www.planetcinema.co.il/il/data-api-service/v1/quickbook/10100'
    $until = (Get-Date).AddDays(60).ToString('yyyy-MM-dd')
    $lastDay = (Get-Date).AddDays(7).ToString('yyyy-MM-dd')
    $pcin = (Get-Text "$pb/cinemas/with-event/until/${until}?attr=&lang=he_IL" | ConvertFrom-Json).body.cinemas
    if (-not $pcin) { throw 'Planet returned no cinemas.' }
    $pfilms = @{}; $pevents = New-Object System.Collections.ArrayList; $pcinemas = @()
    foreach ($c in $pcin) {
        $pcinemas += [pscustomobject]@{ id = [string]$c.id; name = [string]$c.displayName }
        try { $dates = (Get-Text "$pb/dates/in-cinema/$($c.id)/until/${until}?attr=&lang=he_IL" | ConvertFrom-Json).body.dates } catch { $dates = @() }
        foreach ($day in @($dates | Where-Object { $_ -le $lastDay })) {
            try { $b = (Get-Text "$pb/film-events/in-cinema/$($c.id)/at-date/${day}?attr=&lang=he_IL" | ConvertFrom-Json).body } catch { continue }
            foreach ($f in $b.films) {
                if (-not $pfilms.ContainsKey($f.id)) {
                    $pfilms[$f.id] = [pscustomobject]@{ n = $f.name; l = $f.length; p = $f.posterLink; v = $f.videoLink; y = $f.releaseYear; a = ($f.attributeIds -join '|') }
                }
            }
            foreach ($e in $b.events) {
                [void]$pevents.Add([pscustomobject]@{
                    c = [string]$e.cinemaId; f = $e.filmId; t = $e.eventDateTime; d = $e.businessDay; h = $e.auditorium
                    r = $e.availabilityRatio; s = [bool]$e.soldOut; b = $e.bookingRouterLaunchLink; a = ($e.attributeIds -join '|')
                    du = ($e.languages.dubbed -join ','); su = ($e.languages.subtitles -join ',')
                })
            }
        }
    }
    if ($pevents.Count -eq 0) { throw 'Planet returned no screenings.' }
    $pjson = [pscustomobject]@{ cinemas = $pcinemas; films = $pfilms; events = $pevents } | ConvertTo-Json -Depth 4 -Compress
    $html = Set-Block $html 'lev-planet' ("window.PLANET = " + (Safe $pjson) + ";")
}
catch { $planetError = Clean $_.Exception.Message }

Step 'Cinema City...'
# 3c. Cinema City runs on the same ticketing system as Lev. Keep the next 8 days and only
#     the fields the page uses, plus trailers/posters and every hall's seat plan.
$ccError = $null
try {
    Add-Type -AssemblyName System.Web.Extensions
    $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $ser.MaxJsonLength = [int]::MaxValue
    $cc = 'https://tickets.cinema-city.co.il/api'
    $ccLast = (Get-Date).AddDays(7).ToString('yyyy-MM-dd')
    $ccRaw = $ser.DeserializeObject((Get-Text "$cc/presentations?includeSynopsis=0"))
    $keep = 'id','dateTime','businessDate','featureName','featureAdditionalName','featureId','featureRatingName','durationInMinutes',
            'venueId','venueName','venueTinyName','venueLocationId','locationName','languageISO','subbedLanguageISO','dubbedLanguageISO','soldout','availRatio'
    # Small JSON writer (PowerShell's own one trips over the values read above)
    function J($v) {
        if ($null -eq $v) { return 'null' }
        if ($v -is [bool]) { if ($v) { return 'true' } else { return 'false' } }
        if ($v -is [int] -or $v -is [long] -or $v -is [decimal] -or $v -is [double]) { return ([decimal]$v).ToString([Globalization.CultureInfo]::InvariantCulture) }
        $s = [string]$v
        $s = $s.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t').Replace('<', $LT)
        return '"' + $s + '"'
    }
    $ccList = New-Object System.Collections.Generic.List[string]
    $ccPlans = @{}
    foreach ($p in $ccRaw['presentations']) {
        if ($p['private'] -or [string]$p['businessDate'] -gt $ccLast) { continue }
        $ccList.Add('{' + (($keep | ForEach-Object { '"' + $_ + '":' + (J $p[$_]) }) -join ',') + '}')
        $ccPlans[[string]$p['venueId']] = [string]$p['seatplanId']
    }
    if ($ccList.Count -eq 0) { throw 'Cinema City returned no screenings.' }
    $ccFilms = New-Object System.Collections.Generic.List[string]
    foreach ($f in $ser.DeserializeObject((Get-Text "$cc/features"))) {
        $ccFilms.Add((J ([string]$f['id'])) + ':{"t":' + (J $f['trailer']) + ',"p":' + (J $f['imageData']) + '}')
    }
    $ccHalls = @()
    foreach ($v in ($ccPlans.Keys | Sort-Object { [int]$_ })) {
        try {
            $wc = New-Client
            $wc.Headers.Add('Content-Type', 'application/json')
            $s = $wc.UploadString("$cc/seats/seatplanV2?venueId=$v&seatplanId=$($ccPlans[$v])", 'POST', '{}')
            if ($s.TrimStart().StartsWith('{')) { $ccHalls += "`"$v`": " + (Safe ([regex]::Replace($s, '"i":"data:[^"]*"', '"i":null'))) }
        }
        catch { }  # that hall just shows without a seat map
    }
    $html = Set-Block $html 'lev-cc' ("window.CC = {presentations: [" + ($ccList -join ',') +
        "], films: {" + ($ccFilms -join ',') + "}, halls: {" + ($ccHalls -join ', ') + "}};")
}
catch { $ccError = Clean $_.Exception.Message }

Step 'Movieland...'
# 3d. Movieland: the schedule, halls and seat plans come from its ticketing system (biggerpicture),
#     which lets anyone in as a guest; posters and trailers come from movieland.co.il.
$mlError = $null
try {
    if (-not $ser) {
        Add-Type -AssemblyName System.Web.Extensions
        $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $ser.MaxJsonLength = [int]::MaxValue
    }
    if (-not (Get-Command J -ErrorAction SilentlyContinue)) {
        function J($v) {
            if ($null -eq $v) { return 'null' }
            if ($v -is [bool]) { if ($v) { return 'true' } else { return 'false' } }
            if ($v -is [int] -or $v -is [long] -or $v -is [decimal] -or $v -is [double]) { return ([decimal]$v).ToString([Globalization.CultureInfo]::InvariantCulture) }
            $s = [string]$v
            $s = $s.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t').Replace('<', $LT)
            return '"' + $s + '"'
        }
    }
    $bp = 'https://pub-api-use1.biggerpicture.ai/ecomAPI/public/api'
    $mlSite = 'https://movieland.co.il'
    function BP-Token($siteId) {
        $wc = New-Client
        $wc.Headers.Add('Content-Type', 'application/json')
        $wc.Headers.Add('Origin', 'https://ecom.biggerpicture.ai')
        return $ser.DeserializeObject($wc.UploadString("$bp/sys/login", 'POST', '{"siteId":' + $siteId + ',"saleChannelCode":"web","language":"he_IL"}'))['token']
    }
    function BP-Get($path, $token) {
        $wc = New-Client
        $wc.Headers.Add('Authorization', "Bearer $token")
        $wc.Headers.Add('Origin', 'https://ecom.biggerpicture.ai')
        return $ser.DeserializeObject([System.Text.Encoding]::UTF8.GetString($wc.DownloadData("$bp/$path")))
    }
    # Branch ids, read from the Movieland home page (with the known ones as a fallback)
    $mlSites = @()
    try { $mlSites = @([regex]::Matches((Get-Text "$mlSite/"), '"TixTheatreId":(\d+)') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique) } catch {}
    if (-not $mlSites.Count) { $mlSites = @('1290', '1291', '1292', '1293', '1294', '1296') }
    $mlFrom = (Get-Date).ToString('yyyy-MM-dd'); $mlTo = (Get-Date).AddDays(7).ToString('yyyy-MM-dd')
    $LANG = @{ Hebrew = 'he'; English = 'en'; Arabic = 'ar'; Russian = 'ru'; French = 'fr'; Spanish = 'es'; Italian = 'it'; German = 'de'; Japanese = 'ja'; Korean = 'ko' }
    $mlList = New-Object System.Collections.Generic.List[string]
    $mlHalls = New-Object System.Collections.Generic.List[string]
    $movieOfEvent = @{}; $mlMovies = @{}
    foreach ($site in $mlSites) {
        # Movieland's own list links each screening to its film page (poster and trailer)
        try {
            foreach ($m in $ser.DeserializeObject((Get-Text "$mlSite/api/Events?TheatreId=$site&isForSelectedTheaterOnly=true&isHideVODRent=true"))) {
                if (-not $mlMovies.ContainsKey([string]$m['MovieId'])) { $mlMovies[[string]$m['MovieId']] = [string]$m['Pic'] }
                foreach ($d in $m['Dates']) { $movieOfEvent["$site-$($d['EventId'])"] = [string]$m['MovieId'] }
            }
        } catch {}
        try { $token = BP-Token $site } catch { continue }
        try { $events = (BP-Get "cus/eventMaster/0/site/$site/startDate/$mlFrom/endDate/$mlTo" $token)['list'] } catch { continue }
        $plans = @{}
        # Screenings that already started are missing from Movieland's list; use the same film's other screenings
        $movieOfFilm = @{}
        foreach ($e in $events) { $mv = $movieOfEvent["$site-$($e['eventId'])"]; if ($mv) { $movieOfFilm[[string]$e['eventMasterId']] = $mv } }
        foreach ($e in $events) {
            if ($e['isPrivate'] -or -not $e['isActive']) { continue }
            $bd = ([string]$e['businessDate']).Substring(0, 10)
            if ($bd -gt $mlTo) { continue }
            $rating = [string]$e['ratingShortName']
            if ($rating -match '^\+(\d+)$') { $rating = $Matches[1] + '+' }
            elseif ([int]$e['ratingId'] -eq 1) { $rating = [string]$e['ratingName'] }   # all ages
            elseif ($rating -match "[$([char]0x0590)-$([char]0x05FF)]") { $rating = $null }               # "not rated yet" and cut-off texts
            $du = $null; $su = $null
            if ($e['isDubbed']) { $du = $LANG[[string]$e['dubbedLanguageName']] }
            if ($e['isSubtitled'] -and -not $du) { $su = $LANG[[string]$e['subtitledLanguageName']] }
            $fmt = @()
            try { $fmt = @($ser.DeserializeObject([string]$e['eventAttributes']) | ForEach-Object { [string]$_['name'] } | Where-Object { $_ -notin '2D', 'SUBTITLE', 'DUB', 'DUBBED' }) } catch {}
            $cap = [double]$e['capacity']; $avail = [double]$e['seatsAvailable']
            $movie = $movieOfEvent["$site-$($e['eventId'])"]
            if (-not $movie) { $movie = $movieOfFilm[[string]$e['eventMasterId']] }
            $fields = [ordered]@{
                id = [string]$e['eventCode']; dateTime = [string]$e['dateTimeOfEvent']; businessDate = $bd
                featureName = ([string]$e['eventMasterDisplayName']).Trim() -replace '\s+', ' '
                featureAdditionalName = ([string]$e['originalName']).Trim(' ', '|')
                featureId = $(if ($movie) { $movie } else { 'e' + $e['eventMasterId'] })
                featureRatingName = $rating; durationInMinutes = $e['lengthInMinutes']
                venueId = [string]$e['venueId']; venueName = ([string]$e['venueDisplayName']).Trim()
                venueLocationId = [string]$site; locationName = $(if ($e['siteShortName']) { ([string]$e['siteShortName']).Trim() } else { ([string]$e['siteDisplayName']).Trim() })
                dubbedLanguageISO = $du; subbedLanguageISO = $su
                soldout = ($cap -gt 0 -and $avail -le 0)
                availRatio = $(if ($cap -gt 0) { [Math]::Round([Math]::Max(0, $avail) / $cap, 3) } else { $null })
            }
            $mlList.Add('{' + (($fields.Keys | ForEach-Object { '"' + $_ + '":' + (J $fields[$_]) }) -join ',') + ',"featureAttributeAbbrvs":[' + (($fmt | ForEach-Object { J $_ }) -join ',') + ']}')
            $plans[[string]$e['venueId']] = [string]$e['venueSeatPlanId']
        }
        # Seat plans, turned into the same shape as Lev's so the page draws them the same way
        foreach ($v in $plans.Keys) {
            try {
                $secs = @(); $si = 0
                foreach ($sec in (BP-Get "cus/venue/$v/venueSeatPlan/$($plans[$v])" $token)['list']) {
                    $rows = [ordered]@{}
                    foreach ($st in $sec['venueSectionSeats']) {
                        if ($st['isDeleted']) { continue }
                        $rn = [string]$st['rowName']
                        if (-not $rows.Contains($rn)) { $rows[$rn] = New-Object System.Collections.Generic.List[string] }
                        $cx = [Math]::Round(([double]$st['xCoordinate'] + [double]$st['xOffsetPercent'] / 100) * 55)
                        $cy = [Math]::Round(([double]$st['yCoordinate'] + [double]$st['yOffsetPercent'] / 100) * 55)
                        $hc = (',' + [string]$st['venueSeatAttributes'] + ',') -match ',10,'
                        $rows[$rn].Add('"' + $rows[$rn].Count + '":{"n":' + (J ([string]$st['seatName'])) + ',"rd":{"cx":' + $cx + ',"cy":' + $cy + '}' + $(if ($hc) { ',"hc":true' } else { '' }) + '}')
                    }
                    if (-not $rows.Count) { continue }
                    $gx = [Math]::Round([double]$sec['xCoordinate'] * 55); $gy = [Math]::Round([double]$sec['yCoordinate'] * 55)
                    $rparts = $rows.Keys | ForEach-Object { (J $_) + ':{"n":' + (J $_) + ',"S":{' + ($rows[$_] -join ',') + '}}' }
                    $secs += '"' + $si + '":{"n":' + (J ([string]$sec['displayName'])) + ',"G":{"0":{"rd":{"x":' + $gx + ',"y":' + $gy + '},"R":{' + ($rparts -join ',') + '}}}}'
                    $si++
                }
                if ($secs.Count) { $mlHalls.Add("`"${site}:$v`":{`"S`":{" + ($secs -join ',') + '}}') }
            }
            catch { }  # that hall just shows without a seat map
        }
    }
    if ($mlList.Count -eq 0) { throw 'Movieland returned no screenings.' }
    # Posters (resized by Movieland's own server) and trailers from each film's page
    $mlFilms = New-Object System.Collections.Generic.List[string]
    $mlKnown = Old-Trailers $html 'lev-ml'
    $budget = [Diagnostics.Stopwatch]::StartNew()
    foreach ($id in $mlMovies.Keys) {
        $poster = $null; $trailer = $mlKnown[[string]$id]
        if ($mlMovies[$id]) { $poster = "$mlSite/images/" + [Uri]::EscapeDataString($mlMovies[$id]) + '?w=96&h=136&mode=crop' }
        # Only new films' pages are fetched, and never for more than 40 seconds in all
        if (-not $trailer -and $budget.Elapsed.TotalSeconds -lt 40) {
            try {
                $mp = Get-Text "$mlSite/movie/$id"
                $tm = [regex]::Match($mp, '(?:youtube\.com/(?:watch\?v=|embed/)|youtu\.be/)([A-Za-z0-9_-]{11})')
                if ($tm.Success) { $trailer = 'https://www.youtube.com/watch?v=' + $tm.Groups[1].Value }
            } catch {}
        }
        $mlFilms.Add((J $id) + ':{"t":' + (J $trailer) + ',"p":' + (J $poster) + '}')
    }
    $html = Set-Block $html 'lev-ml' ("window.ML = {presentations: [" + ($mlList -join ',') +
        "], films: {" + ($mlFilms -join ',') + "}, halls: {" + ($mlHalls -join ',') + "}};")
}
catch { $mlError = Clean $_.Exception.Message }

# Cinematheque fallback for GitHub, whose servers the ticket system refuses: the screenings come from
# cinema.co.il's day calendar, while seat plans and known films' details are kept from the PC's last update.
# The calendar has no seat counts, so "how full" is left out rather than shown from an older update.
function Update-CinFromWeb([string]$html) {
    # The previous Cinematheque data in the page
    $s = $html.IndexOf('<script id="lev-cin">')
    $block = $html.Substring($s, $html.IndexOf('</script>', $s) - $s)
    $prevShows = @(); $prevFilms = ''; $prevHalls = ''
    $a = $block.IndexOf('presentations: ['); $b = $block.IndexOf('], films: {'); $c = $block.IndexOf('}, halls: {'); $d = $block.LastIndexOf('}};')
    if ($a -ge 0 -and $b -gt $a -and $c -gt $b -and $d -gt $c) {
        $prevShows = @($ser.DeserializeObject('[' + $block.Substring($a + 16, $b - $a - 16) + ']'))
        $prevFilms = $block.Substring($b + 11, $c - $b - 11)
        $prevHalls = $block.Substring($c + 11, $d - $c - 11)
    }
    $filmOfId = @{}; $filmOfTitle = @{}; $venueOfHall = @{}; $loc = '1'; $locName = $null
    foreach ($p in $prevShows) {
        $filmOfId[[string]$p['id']] = [string]$p['featureId']
        $filmOfTitle[([string]$p['featureName']).Trim()] = [string]$p['featureId']
        $venueOfHall[([string]$p['venueName']).Trim()] = [string]$p['venueId']
        if ($p['venueLocationId']) { $loc = [string]$p['venueLocationId'] }
        if ($p['locationName']) { $locName = [string]$p['locationName'] }
    }
    $calUrl = 'https://www.cinema.co.il/wp-content/themes/cinematheque_new_theme/ajax_data.php'
    $list = New-Object System.Collections.Generic.List[string]
    $newFilms = [ordered]@{}; $seen = @{}
    for ($i = 0; $i -le 7; $i++) {
        $day = (Get-Date).AddDays($i)
        $wc = New-Client
        $wc.Headers.Add('Content-Type', 'application/x-www-form-urlencoded')
        $cal = $wc.UploadString("${calUrl}?action=get_cal_data&ids=&date=$($day.ToString('yyyy-MM-dd'))&lang=he&tc=$i", 'POST', '')
        foreach ($entry in ($cal -split 'class="calentry"')) {
            $om = [regex]::Match($entry, 'cintlv\.pres\.global/order/(\d+)')
            $tm = [regex]::Match($entry, 'class="item_center">\s*(\d{1,2}):(\d{2})<span class="ven"><br/>([^<]+)</span>')
            $nm = [regex]::Match($entry, 'class="s1">\s*([^<]*?)\s*</span>')
            if (-not ($om.Success -and $tm.Success -and $nm.Success)) { continue }
            $id = $om.Groups[1].Value
            if ($seen[$id]) { continue }
            $seen[$id] = 1
            $title = [System.Net.WebUtility]::HtmlDecode($nm.Groups[1].Value).Trim()
            $hall = [System.Net.WebUtility]::HtmlDecode($tm.Groups[3].Value).Trim()
            # Screenings after midnight belong to the evening before
            $start = $day.Date.AddHours([int]$tm.Groups[1].Value).AddMinutes([int]$tm.Groups[2].Value)
            if ([int]$tm.Groups[1].Value -lt 5) { $start = $start.AddDays(1) }
            $len = $null
            $lm = [regex]::Match($entry, 'class="event_free_txt">\s*(\d+)')
            if ($lm.Success) { $len = [int]$lm.Groups[1].Value }
            $fid = $filmOfId[$id]
            if (-not $fid) { $fid = $filmOfTitle[$title] }
            if (-not $fid) {
                if (-not $newFilms.Contains($title)) {
                    $im = [regex]::Match($entry, 'class="evn_right"><img src="(https://www\.cinema\.co\.il/[^"]+)"')
                    $pm = [regex]::Match($entry, 'href="(https://www\.cinema\.co\.il/event/[^"]+)"')
                    $newFilms[$title] = @{ id = 'w' + ($newFilms.Count + 1); p = $(if ($im.Success) { [Uri]::EscapeUriString([Uri]::UnescapeDataString($im.Groups[1].Value)) }); u = $(if ($pm.Success) { $pm.Groups[1].Value }) }
                }
                $fid = $newFilms[$title].id
            }
            $venue = $venueOfHall[$hall]
            if (-not $venue) { $venue = 'w-' + ($hall -replace '\D', '') }
            $f = [ordered]@{
                id = $id; dateTime = $start.ToString('yyyy-MM-dd HH:mm'); businessDate = $day.ToString('yyyy-MM-dd')
                featureName = $title; featureAdditionalName = ''; featureId = $fid; featureRatingName = $null; durationInMinutes = $len
                venueId = $venue; venueName = $hall; venueTinyName = ($hall -replace '\D', ''); venueLocationId = $loc; locationName = $locName
                languageISO = $null; subbedLanguageISO = $null; dubbedLanguageISO = $null; soldout = 0; availRatio = $null
            }
            $list.Add('{' + (($f.Keys | ForEach-Object { '"' + $_ + '":' + (J $f[$_]) }) -join ',') + '}')
        }
    }
    if ($list.Count -eq 0) { throw 'cinema.co.il calendar had no screenings.' }
    # New films (not in the PC's last update): poster from the calendar, trailer from the film's page
    $known = Old-Trailers $html 'lev-cin'
    $films = @()
    foreach ($t in $newFilms.Keys) {
        $nf = $newFilms[$t]; $tr = $known[$nf.id]
        if (-not $tr -and $nf.u) {
            try {
                $ym = [regex]::Match((Get-Text $nf.u), '(?:youtube\.com/(?:watch\?v=|embed/)|youtu\.be/)([A-Za-z0-9_-]{11})')
                if ($ym.Success) { $tr = 'https://www.youtube.com/watch?v=' + $ym.Groups[1].Value }
            } catch {}
        }
        $films += (J $nf.id) + ':{"t":' + (J $tr) + ',"p":' + (J $nf.p) + ',"u":' + (J $nf.u) + '}'
    }
    $allFilms = @($prevFilms.Trim(), ($films -join ',')) | Where-Object { $_ }
    return Set-Block $html 'lev-cin' ("window.CIN_FETCHED = '$stamp'; window.CIN = {presentations: [" + ($list -join ',') +
        "], films: {" + ($allFilms -join ',') + "}, halls: {" + $prevHalls + "}};")
}

Step 'Cinematheque...'
# 3e. Tel Aviv Cinematheque: same ticketing system as Lev and Cinema City for screenings and seat plans.
#     Its ticket system has few posters and no trailers, so those come from cinema.co.il's own calendar and film pages.
$cinError = $null
try {
    $cinFrom = (Get-Date).ToString('yyyy-MM-dd'); $cinTo = (Get-Date).AddDays(7).ToString('yyyy-MM-dd')
    # cintlv.pres.global forwards to cintlv.presglobal.store; try that one directly first, as the
    # forwarding address refuses some servers (e.g. GitHub's)
    if ($env:CIN_TEST_WEB -eq '1') { throw 'test: pretending the ticket system refused us' }
    $cinText = $null; $cinTried = @()
    foreach ($cb in 'https://cintlv.presglobal.store/api', 'https://cintlv.pres.global/api') {
        try { $cinText = Get-Text "$cb/presentations?includeSynopsis=0"; break }
        catch { $cinTried += ($cb -replace '^https://([^/]+)/api$', '$1') + ': ' + $_.Exception.Message }
    }
    if (-not $cinText) { throw ($cinTried -join ' / ') }
    $cinRaw = $ser.DeserializeObject($cinText)
    $keep = 'id','dateTime','businessDate','featureName','featureAdditionalName','featureId','featureRatingName','durationInMinutes',
            'venueId','venueName','venueTinyName','venueLocationId','locationName','languageISO','subbedLanguageISO','dubbedLanguageISO','soldout','availRatio'
    $cinList = New-Object System.Collections.Generic.List[string]
    $cinPlans = @{}; $filmOfShow = @{}; $cinFilmIds = @{}
    foreach ($p in $cinRaw['presentations']) {
        $bd = [string]$p['businessDate']
        if ($p['private'] -or $bd -lt $cinFrom -or $bd -gt $cinTo) { continue }
        if ([int]$p['venueTypeId'] -eq 101 -or [string]$p['venueName'] -eq 'VOD') { continue }   # streaming, not a hall
        $cinList.Add('{' + (($keep | ForEach-Object { '"' + $_ + '":' + (J $p[$_]) }) -join ',') + '}')
        $cinPlans[[string]$p['venueId']] = [string]$p['seatplanId']
        $filmOfShow[[string]$p['id']] = [string]$p['featureId']
        $cinFilmIds[[string]$p['featureId']] = 1
    }
    if ($cinList.Count -eq 0) { throw 'The Cinematheque returned no screenings.' }
    $cinPoster = @{}; $cinTrailer = @{}; $cinPage = @{}
    foreach ($f in $ser.DeserializeObject((Get-Text "$cb/features"))) {
        $id = [string]$f['id']
        if (-not $cinFilmIds.ContainsKey($id)) { continue }
        if ([string]$f['imageData'] -match '^https://') { $cinPoster[$id] = [string]$f['imageData'] }
        if ([string]$f['trailer'] -match '^https://') { $cinTrailer[$id] = ([string]$f['trailer']).Trim() }
    }
    # cinema.co.il's day calendar: each screening's ticket number, poster and film page
    $calUrl = 'https://www.cinema.co.il/wp-content/themes/cinematheque_new_theme/ajax_data.php'
    for ($i = 0; $i -le 7; $i++) {
        $day = (Get-Date).AddDays($i).ToString('yyyy-MM-dd')
        try {
            $wc = New-Client
            $wc.Headers.Add('Content-Type', 'application/x-www-form-urlencoded')
            $cal = $wc.UploadString("${calUrl}?action=get_cal_data&ids=&date=$day&lang=he&tc=$i", 'POST', '')
        } catch { continue }
        foreach ($entry in ($cal -split 'class="calentry"')) {
            $om = [regex]::Match($entry, 'cintlv\.pres\.global/order/(\d+)')
            if (-not $om.Success) { continue }
            $fid = $filmOfShow[$om.Groups[1].Value]
            if (-not $fid) { continue }
            $im = [regex]::Match($entry, 'class="evn_right"><img src="(https://www\.cinema\.co\.il/[^"]+)"')
            if ($im.Success -and -not $cinPoster[$fid]) { $cinPoster[$fid] = $im.Groups[1].Value }
            $pm = [regex]::Match($entry, 'href="(https://www\.cinema\.co\.il/event/[^"]+)"')
            if ($pm.Success -and -not $cinPage[$fid]) { $cinPage[$fid] = $pm.Groups[1].Value }
        }
    }
    # Trailers from each film's page on cinema.co.il
    $cinKnown = Old-Trailers $html 'lev-cin'
    $budget = [Diagnostics.Stopwatch]::StartNew()
    foreach ($fid in @($cinPage.Keys)) {
        if ($cinTrailer[$fid]) { continue }
        if ($cinKnown[$fid]) { $cinTrailer[$fid] = $cinKnown[$fid]; continue }   # found on an earlier run
        if ($budget.Elapsed.TotalSeconds -gt 40) { continue }
        try {
            $tm = [regex]::Match((Get-Text $cinPage[$fid]), '(?:youtube\.com/(?:watch\?v=|embed/)|youtu\.be/)([A-Za-z0-9_-]{11})')
            if ($tm.Success) { $cinTrailer[$fid] = 'https://www.youtube.com/watch?v=' + $tm.Groups[1].Value }
        } catch {}
    }
    $cinFilms = New-Object System.Collections.Generic.List[string]
    foreach ($fid in $cinFilmIds.Keys) {
        $poster = $cinPoster[$fid]
        if ($poster) { $poster = [Uri]::EscapeUriString([Uri]::UnescapeDataString($poster)) }   # Hebrew file names, encoded once
        $cinFilms.Add((J $fid) + ':{"t":' + (J $cinTrailer[$fid]) + ',"p":' + (J $poster) + ',"u":' + (J $cinPage[$fid]) + '}')
    }
    # Seat plans: this site only gives the older grid format (row, seat, x, y), turned into Lev's shape here
    $cinHalls = @()
    foreach ($v in ($cinPlans.Keys | Sort-Object { [int]$_ })) {
        try {
            $sp = $ser.DeserializeObject((Get-Text "$cb/seats/seatplan?venueId=$v&seatplanId=$($cinPlans[$v])"))
            $secs = @(); $yShift = 0
            foreach ($sk in $sp['sections'].Keys) {
                $sec = $sp['sections'][$sk]; $rparts = @(); $maxY = 0
                foreach ($g in $sec['groups'].Values) {
                    foreach ($row in $g['rows'].Values) {
                        $seats = @(); $n = 0
                        foreach ($st in $row['seats'].Values) {
                            $x = [double]$st['XCoordinate'] + [double]$st['XOffsetPercent'] / 100
                            $y = [double]$st['YCoordinate'] + [double]$st['YOffsetPercent'] / 100
                            if ($y -gt $maxY) { $maxY = $y }
                            $seats += '"' + ($n++) + '":{"n":' + (J ([string]$st['SeatName'])) + ',"rd":{"cx":' + [Math]::Round($x * 55) + ',"cy":' + [Math]::Round(($y + $yShift) * 55) + '}}'
                        }
                        if ($seats.Count) { $rparts += (J ([string]$row['RowName'] + "#$sk")) + ':{"n":' + (J ([string]$row['RowName'])) + ',"S":{' + ($seats -join ',') + '}}' }
                    }
                }
                if ($rparts.Count) {
                    $secs += (J ([string]$sk)) + ':{"n":' + (J ([string]$sec['SectionName'])) + ',"G":{"0":{"rd":{"x":0,"y":0},"R":{' + ($rparts -join ',') + '}}}}'
                    $yShift += $maxY + 2   # a further section (e.g. a balcony) goes below the one before
                }
            }
            if ($secs.Count) { $cinHalls += "`"$v`":{`"S`":{" + ($secs -join ',') + '}}' }
        }
        catch { }  # that hall just shows without a seat map
    }
    $html = Set-Block $html 'lev-cin' ("window.CIN_FETCHED = '$stamp'; window.CIN = {presentations: [" + ($cinList -join ',') +
        "], films: {" + ($cinFilms -join ',') + "}, halls: {" + ($cinHalls -join ', ') + "}};")
}
catch {
    $cinError = Clean $_.Exception.Message
    # The Cinematheque's ticket site refuses GitHub's servers; there the PC's last schedule is kept
    # and the page only warns once that schedule gets old (it knows when it was downloaded).
    if ($env:GITHUB_ACTIONS -eq 'true' -or $env:CIN_TEST_WEB -eq '1') {
        $note = "$stamp ticket system: $cinError"
        try { $html = Update-CinFromWeb $html; $note += ' | cinema.co.il calendar: OK' }
        catch { $note += ' | cinema.co.il calendar: ' + (Clean $_.Exception.Message) }
        [System.IO.File]::WriteAllText((Join-Path $here 'cinematheque-error.txt'), $note, $utf8)
        $cinError = $null
    }
}

Step 'streaming posters...'
# 4. Posters for Lev's streaming titles (Lev's film list has none for them).
#    Each film is looked up once and remembered inside lev.html, so later runs only check new titles.
#    Older films come back as a whole embedded picture; those are shrunk to a small thumbnail.
#    Remembered values: a web address, a thumbnail, or "-" for a film with no usable poster.
Add-Type -AssemblyName System.Drawing
function Get-Thumb([string]$dataUri) {
    $bytes = [Convert]::FromBase64String($dataUri.Substring($dataUri.IndexOf(',') + 1))
    $ms = New-Object IO.MemoryStream(, $bytes)
    $src = [System.Drawing.Image]::FromStream($ms)
    try {
        $w = 80; $h = 114   # poster shape; about 1.5x the size it is shown at
        $scale = [Math]::Max($w / $src.Width, $h / $src.Height)
        $dw = $src.Width * $scale; $dh = $src.Height * $scale
        $bmp = New-Object System.Drawing.Bitmap($w, $h)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.DrawImage($src, [single](($w - $dw) / 2), [single](($h - $dh) / 2), [single]$dw, [single]$dh)
        $g.Dispose()
        $jpeg = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
        $ep = New-Object System.Drawing.Imaging.EncoderParameters(1)
        $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]62)
        $out = New-Object IO.MemoryStream
        $bmp.Save($out, $jpeg, $ep)
        $bmp.Dispose()
        return 'data:image/jpeg;base64,' + [Convert]::ToBase64String($out.ToArray())
    }
    finally { $src.Dispose(); $ms.Dispose() }
}
if ($json) {
    try {
        $cache = @{}
        $pi = $html.IndexOf('<script id="lev-posters">')
        if ($pi -ge 0) {
            $pe = $html.IndexOf('</script>', $pi)
            foreach ($m in [regex]::Matches($html.Substring($pi, $pe - $pi), '"(\d+)":"([^"]*)"')) { $cache[$m.Groups[1].Value] = $m.Groups[2].Value }
        }
        $ids = @{}
        foreach ($m in [regex]::Matches($json, '"featureId":\s*(\d+),[^{}]*?"venueLocationId":\s*1,')) { $ids[$m.Groups[1].Value] = 1 }
        $looked = 0
        foreach ($id in @($ids.Keys)) {
            if (($cache.ContainsKey($id) -and $cache[$id] -ne '') -or $looked -ge 400) { continue }
            $looked++
            try {
                $r = (Get-Text "$api/features/$id/image").Trim()
                if ($r -match '^"(https://[A-Za-z0-9./_\-]+)"$') { $cache[$id] = $Matches[1] }
                elseif ($r.StartsWith('"data:image/')) {
                    try { $cache[$id] = Get-Thumb ($r.Trim('"').Replace('\/', '/')) } catch { $cache[$id] = '-' }
                }
                else { $cache[$id] = '-' }
            }
            catch { }  # not remembered, so it is tried again next time
        }
        $pairs2 = foreach ($k in ($cache.Keys | Sort-Object { [int]$_ })) { "`"$k`":`"$($cache[$k])`"" }
        $html = Set-Block $html 'lev-posters' ("window.LEV_POSTERS = {" + ($pairs2 -join ',') + "};")
    }
    catch { }  # posters are optional; the page shows a blank frame instead
}

Step 'saving the page...'
# 5. Status shown on the page
function JsStr($s) { if ($s) { return "'$s'" } else { return 'null' } }
$html = Set-Block $html 'lev-status' ("window.LEV_ERROR = $(JsStr $scheduleError); window.LEV_HALLS_ERROR = $(JsStr $hallsError); window.LEV_FILMS_ERROR = $(JsStr $filmsError); window.PLANET_ERROR = $(JsStr $planetError); window.CC_ERROR = $(JsStr $ccError); window.ML_ERROR = $(JsStr $mlError); window.CIN_ERROR = $(JsStr $cinError); window.LEV_ERROR_AT = '$stamp';")

$tmp = "$page.tmp"
[System.IO.File]::WriteAllText($tmp, $html, $utf8)
Move-Item -LiteralPath $tmp -Destination $page -Force

Step 'uploading to the phone page...'
# 6. Upload the page to the phone web address (GitHub Pages), if that is set up.
#    One commit is kept and replaced each time, so the repository stays small.
$publishDir = Join-Path $env:LOCALAPPDATA 'LevPublish'
if (Test-Path (Join-Path $publishDir '.git')) {
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $log = @("$stamp")
    try {
        Copy-Item -LiteralPath $page -Destination (Join-Path $publishDir 'index.html') -Force
        # The updater and GitHub's run settings go along too, so GitHub can run the same update (from the phone, or on its schedule)
        Copy-Item -LiteralPath (Join-Path $here 'update.ps1') -Destination (Join-Path $publishDir 'update.ps1') -Force
        $wf = Join-Path $here 'github-workflow.yml'
        if (Test-Path -LiteralPath $wf) {
            New-Item -ItemType Directory -Force (Join-Path $publishDir '.github\workflows') | Out-Null
            Copy-Item -LiteralPath $wf -Destination (Join-Path $publishDir '.github\workflows\update.yml') -Force
        }
        $log += (& git -C $publishDir add -A 2>&1 | Out-String)
        & git -C $publishDir rev-parse --verify --quiet HEAD *> $null
        if ($LASTEXITCODE -eq 0) {
            $log += (& git -C $publishDir commit --amend --no-edit --date=now --quiet 2>&1 | Out-String)
        } else {
            $log += (& git -C $publishDir commit -m 'Lev Tel Aviv schedule' --quiet 2>&1 | Out-String)
        }
        # Always use this account, and never wait for a sign-in window in the background
        & git -C $publishDir config credential.username yonatanbendavid 2>&1 | Out-Null
        # The repository was renamed from "lev"; the phone address is now /mabakolnoa
        & git -C $publishDir remote set-url origin 'https://github.com/yonatanbendavid/mabakolnoa.git' 2>&1 | Out-Null
        $env:GCM_INTERACTIVE = 'never'
        $log += (& git -C $publishDir push --force --quiet origin HEAD:main 2>&1 | Out-String)
        $log[0] += $(if ($LASTEXITCODE -eq 0) { ' published' } else { ' PUBLISH FAILED' })
    }
    catch { $log[0] += ' PUBLISH FAILED'; $log += $_.Exception.Message }
    finally { $ErrorActionPreference = $prev }
    [System.IO.File]::WriteAllText((Join-Path $here 'publish.log'), ($log -join "`r`n"), $utf8)
}

Step 'Done.'
if (-not $NoOpen) { Start-Process $page }
