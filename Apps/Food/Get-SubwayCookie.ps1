$UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.75 Safari/537.36"
$StoreCode = ""


$SubwayGlobal = Invoke-WebRequest -Uri "https://global.subway.com" -UserAgent $UserAgent
$SubwayListens = Invoke-WebRequest -Uri "https://subwaylistens.com" -UserAgent $UserAgent