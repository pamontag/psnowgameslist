# Install the module on demand
If (-not (Get-Module -ErrorAction Ignore -ListAvailable PowerHTML)) {
    Write-Verbose "Installing PowerHTML module for the current user..."
    Install-Module PowerHTML -ErrorAction Stop
  }
Import-Module -ErrorAction Stop PowerHTML
$html = ConvertFrom-Html  -Uri https://www.playstation.com/en-gb/explore/playstation-now/ps-now-games/#ps-now-games-content-par-section-121120782-co

$outputFile = "PSNOW_CSV.csv"
if(Test-Path -Path $outputFile) {
    Remove-Item $outputFile
}
Add-Content -Path $outputFile -Value '"Gioco","Piattaforma","VotoMetacritic","UserScore","Genre","Players"'
# Find a specific table by its column names, using an XPath
# query to iterate over all tables.
# #ps-now-games-content-par-section-121120782-co-page-section-par-inlinetabs-content-1
for($j=13; $j -le 14; $j++){
    for($i=1;$i -le 13;$i++){
        $letter = $html.SelectSingleNode("//*[contains(@class, 'content tablen-$j  content-$i')]")
        $letter.SelectNodes(".//*[@class='column-cell']") | foreach {
        
            $category = $_.SelectSingleNode(".//h3").InnerText
            if($category -eq "") {
                $category = "PS4"
            } else {
                $category = $category.Trim()
            }

            $games = $_.SelectSingleNode(".//*[@class='richtext  default counter-continue']//p").InnerText
            
            $games = $games.Split([Environment]::NewLine)
            Write-Host $category 
            foreach($game in $games) {
                if($game.Trim() -eq "") {
                    continue
                }
                $game = $game -Replace "&amp;", "&"
                $game = $game -Replace "&nbsp;", ""
                $game = $game -Replace ",", ""
                
                if($game -like "*ANOMALY WARZONE EARTHAPE ESCAPE 2*") {
                    $game = $game -Replace "ANOMALY WARZONE EARTH", ""
                }
                if($game -like "*ALEX KIDD IN THE MIRACLE WORLD*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ANNA EXTENDED EDITION*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ALIEN RAGE*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ALIEN SPIDY*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ALTERED BEAST*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ANARCHY: RUSH HOUR*" -and $category -eq "PS4") {
                    continue
                } 
                if($game -like "*AQUA PANIC!*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*AR NOSURGE: ODE TO AN UNBORN STAR*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ARCANA HEART 3*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ARCANIA THE COMPLETE TALE*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ARMAGEDDON RIDERS*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ASURA'S WRATH*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ATELIER AYESHA: THE ALCHEMIST OF DUSK*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ATELIER ESCHA & LOGY - ALCHEMISTS OF THE DUST SKY*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ATELIER RORONA PLUS: THE ALCHEMIST OF ARLAND*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*ATELIER SHALLIE - ALCHEMISTS OF THE DUSK SEA*" -and $category -eq "PS4") {
                    continue
                }
                if($game -like "*DARK CHRONICLE*") {
                    continue
                }
                if($game -like "DUKE NUKEM FOREVERDYNASTY WARRIORS 6"){
                    Write-Host $game.Trim()                  
                    Add-Content -Path $outputFile -Value "DUKE NUKEM FOREVER,$($category),,,,"
                    Add-Content -Path $outputFile -Value "DYNASTY WARRIORS 6,$($category),,,,"
                    continue                
                }

                if($game -match "^EVERYBODY'S GOLF$"){                    
                    Add-Content -Path $outputFile -Value "$($game.Trim()),PS4,,,"
                    continue
                } 
                if($game -like "*ICO CLASSICS HD*"){
                    $game = "ICO"
                    Add-Content -Path $outputFile -Value "$($game.Trim()),PS2,,,"
                    continue
                }
                if($game -like "*GOD OF WAR HD*"){
                    $game = "GOD OF WAR"
                    Add-Content -Path $outputFile -Value "$($game.Trim()),PS2,,,"
                    continue
                }
                if($game -like "*GOD OF WAR II HD*"){
                    $game = "GOD OF WAR II"
                    Add-Content -Path $outputFile -Value "$($game.Trim()),PS2,,,"
                    continue
                }
                if($game -like "*SHADOW OF THE COLOSSUS*"){
                    $game = "SHADOW OF THE COLOSSUS"
                    Add-Content -Path $outputFile -Value "$($game.Trim()),PS2,,,"
                    continue
                }

                Write-Host $game.Trim()
                if($game -Like "*(PS2 Classics)*") {
                    $game = $game -Replace "\(PS2 Classics\)", ""
                    Add-Content -Path $outputFile -Value "$($game.Trim()),PS2,,,"
                } else {
                    Add-Content -Path $outputFile -Value "$($game.Trim()),$($category),,,,"
                }
            }    
        }
    }
}
 