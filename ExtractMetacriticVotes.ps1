If (-not (Get-Module -ErrorAction Ignore -ListAvailable PowerHTML)) {
    Write-Verbose "Installing PowerHTML module for the current user..."
    Install-Module PowerHTML -ErrorAction Stop
  }
Import-Module -ErrorAction Stop PowerHTML
$games = Import-Csv -Path PSNOW_CSV.csv
$outputFile = "PSNOW_CSV_VOTES.csv"
$errorLogFile = "error.log"
if(Test-Path -Path $outputFile) {
    Remove-Item $outputFile
}
if(Test-Path -Path $errorLogFile) {
    Remove-Item $errorLogFile
}
Add-Content -Path $outputFile -Value '"Gioco","Piattaforma","VotoMetacritic","UserScore","Genre","Players"'

foreach($game in $games) {
    Write-Host "$($game.Gioco) - $($game.Piattaforma)"
    if($($game.VotoMetacritic -eq "FAILED" -or $game.VotoMetacritic -eq "")){
        $gamename = $game.Gioco
        #FIX COMPLETELY WRONG TITLE NAMES
        $gamename = $gamename.Replace("ATELIER ESCHA & LOGY - ALCHEMISTS OF THE DUST SKY","ATELIER ESCHA LOGY ALCHEMISTS OF THE DUSK SKY")
        $gamename = $gamename.Replace("ATELIER SHALLIE - ALCHEMISTS OF THE DUSK SEA","ATELIER SHALLIE ALCHEMISTS OF THE DUSK SEA")
        $gamename = $gamename.Replace("BATMAN - THE TELLTALE SERIES - EPISODE 1: REALM OF SHADOWS","BATMAN THE TELLTALE SERIES   EPISODE 1 REALM OF SHADOWS")
        $gamename = $gamename.Replace("BIG SKY:INFINITY","BIG SKY INFINITY")     
        $gamename = $gamename.Replace("AGATHA CHRISTIE - THE ABC MURDERS", "AGATHA CHRISTIES THE ABC MURDERS")
        $gamename = $gamename.Replace("BORDERLANDS 2 ULTIMATE EDITION", "BORDERLANDS 2")
        $gamename = $gamename.Replace("BORDERLANDS: THE PRE-SEQUEL ULTIMATE EDITION", "BORDERLANDS THE PRE SEQUEL")
        $gamename = $gamename.Replace("CITIES: SKYLINES","CITIES: SKYLINES   PLAYSTATION 4 EDITION")
        $gamename = $gamename.Replace("CARS MATER-NATIONAL","CARS MATER NATIONAL CHAMPIONSHIP")
        $gamename = $gamename.Replace("CASTLEVANIA: LORDS OF SHADOW - MIRROR OF FATE HD","CASTLEVANIA: LORDS OF SHADOW   MIRROR OF FATE HD") 
        $gamename = $gamename.Replace("DEAD ISLAND DEFINITIVE EDITION","DEAD ISLAND DEFINITIVE COLLECTION")
        $gamename = $gamename.Replace("DARK MIST","DARK VOID")
        $gamename = $gamename.Replace("DEADLY PREMONITION: DIRECTORS CUT","DEADLY PREMONITION: THE DIRECTORS CUT")
        $gamename = $gamename.Replace("DISNEY PIXAR BRAVE","BRAVE THE VIDEO GAME")
        $gamename = $gamename.Replace("DISNEY UNIVERSE ULTIMATE EDITION","DISNEY UNIVERSE")
        $gamename = $gamename.Replace("EVE: VALKYRIE - WARZONE","EVE VALKYRIE   WARZONE")
        $gamename = $gamename.Replace("EVERYBODY'S TENNIS","HOT SHOTS TENNIS")
        $gamename = $gamename.Replace("EVERYBODY'S GOLF: WORLD TOUR","HOT SHOTS GOLF OUT OF BOUNDS")
        $gamename = $gamename.Replace("FORBIDDEN SIREN","SIREN")
        $gamename = $gamename.Replace("F.E.A.R. FIRST ENCOUNTER ASSAULT RECON","FEAR")
        $gamename = $gamename.Replace("FAERY : LEGENDS OF AVALON","FAERY LEGENDS OF AVALON")
        $gamename = $gamename.Replace("GIANA SISTERS: TWISTED DREAMS - DIRECTOR'S CUT","GIANA SISTERS TWISTED DREAMS   DIRECTORS CUT")
        $gamename = $gamename.Replace("HUNTER'S TROPHY 2 - AMERICA","HUNTER'S TROPHY 2   AMERICA")
        $gamename = $gamename.Replace("HUNTER'S TROPHY 2 - AUSTRALIA","HUNTER'S TROPHY 2   AUSTRALIA")
        $gamename = $gamename.Replace("JOE DEVER'S LONE WOLF","JOE DEVER'S LONE WOLF CONSOLE EDITION")
        $gamename = $gamename -replace "^KILLZONE$","KILLZONE HD"
        $gamename = $gamename.Replace("LEGO® DISNEY PIRATES OF THE CARIBBEAN THE VIDEOGAME","LEGO PIRATES OF THE CARIBBEAN THE VIDEO GAME")
        $gamename = $gamename -replace "^LOST PLANET$","LOST PLANET EXTREME CONDITION"
        $gamename = $gamename.Replace("METAL GEAR SOLID GROUND ZEROES","METAL GEAR SOLID V GROUND ZEROES")
        $gamename = $gamename.Replace("MONSTER ENERGY SUPERCROSS","MONSTER ENERGY SUPERCROSS   THE OFFICIAL VIDEOGAME")
        $gamename = $gamename.Replace("MUDRUNNER", "SPINTIRES MUDRUNNER")
        $gamename = $gamename.Replace("MXGP3 – THE OFFICIAL MOTOCROSS VIDEOGAME", "MXGP3 THE OFFICIAL MOTOCROSS VIDEOGAME")
        $gamename = $gamename.Replace("MAHJONG TALES : ANCIENT WISDOM", "MAHJONG TALES ANCIENT WISDOM")
        $gamename = $gamename.Replace("MONKEY ISLAND 2 SPECIAL EDITION: LECHUCK’S REVENGE", "MONKEY ISLAND 2 SPECIAL EDITION LECHUCKS REVENGE")
        $gamename = $gamename.Replace("MXGP – THE OFFICIAL MOTOCROSS VIDEOGAME","MXGP THE OFFICIAL MOTOCROSS VIDEOGAME")
        $gamename = $gamename.Replace("NOBUNAGA'S AMBITION : SPHERE OF INFLUENCE","NOBUNAGA'S AMBITION SPHERE OF INFLUENCE")
        $gamename = $gamename.Replace("RISEN 3: TITAN LORDS - ENHANCED EDITION","RISEN 3: TITAN LORDS   ENHANCED EDITION")
        $gamename = $gamename.Replace("RATCHET & CLANK NEXUS","RATCHET CLANK INTO THE NEXUS")
        $gamename = $gamename.Replace("RATCHET & CLANK Q-FORCE","RATCHET CLANK FULL FRONTAL ASSAULT")
        $gamename = $gamename.Replace("RATCHET & CLANK: A CRACK IN TIME","RATCHET CLANK FUTURE A CRACK IN TIME")
        $gamename = $gamename.Replace("RATCHET & CLANK:QUEST FOR BOOTY","RATCHET CLANK FUTURE QUEST FOR BOOTY")
        $gamename = $gamename.Replace("RED JOHNSON'S CHRONICLES - ONE AGAINST ALL","RED JOHNSON'S CHRONICLES   ONE AGAINST ALL")
        $gamename = $gamename.Replace("RESIDENT EVIL CODE:VERONICA X","RESIDENT EVIL CODE VERONICA X HD")
        $gamename = $gamename.Replace("RUNNER 2: A FLAT OUT RUN OF THE RHYTHM ALIEN","bittrip presentsrunner2 future legend of rhythm alien")
        $gamename = $gamename.Replace("SAM & MAX BTS: EPISODE 1 - ICE STATION SANTA","SAM MAX BEYOND TIME AND SPACE")
        $gamename = $gamename.Replace("SAM & MAX BTS: EPISODE 2 - MOAI BETTER BLUES","SAM MAX BEYOND TIME AND SPACE")
        $gamename = $gamename.Replace("SAM & MAX BTS: EPISODE 3 - NIGHT OF THE RAVING DEAD","SAM MAX BEYOND TIME AND SPACE")
        $gamename = $gamename.Replace("SAM & MAX BTS: EPISODE 4 - CHARIOTS OF THE DOGS","SAM MAX BEYOND TIME AND SPACE")
        $gamename = $gamename.Replace("SAM & MAX BTS: EPISODE 5 - WHAT'S NEW BEELZEBUB?","SAM MAX BEYOND TIME AND SPACE")
        $gamename = $gamename.Replace("SAM & MAX THE DEVIL'S PLAYHOUSE - EP 1","sam max the devils playhouse   episode 1 the penal zone")
        $gamename = $gamename.Replace("SAM & MAX THE DEVIL'S PLAYHOUSE - EP 2","sam max the devils playhouse   episode 2 the tomb of sammun mak")
        $gamename = $gamename.Replace("SAM & MAX THE DEVIL'S PLAYHOUSE - EP 3","sam max the devils playhouse   episode 3 they stole maxs brain!")
        $gamename = $gamename.Replace("SAM & MAX THE DEVIL'S PLAYHOUSE - EP 4","sam max the devils playhouse   episode 4 beyond the alley of the dolls")
        $gamename = $gamename.Replace("SAM & MAX THE DEVIL'S PLAYHOUSE - EP 5","sam max the devils playhouse   episode 5 the city that dares not sleep")  
        $gamename = $gamename.Replace("STRONG BAD'S COOL GAME FOR ATTRACTIVE PEOPLE - EPISODE 1","STRONG BAD'S COOL GAME FOR ATTRACTIVE PEOPLE EPISODE 1 HOMESTAR RUINER")
        $gamename = $gamename.Replace("STRONG BAD'S COOL GAME FOR ATTRACTIVE PEOPLE - EPISODE 2","STRONG BAD'S COOL GAME FOR ATTRACTIVE PEOPLE EPISODE 1 HOMESTAR RUINER")
        $gamename = $gamename.Replace("STRONG BAD'S COOL GAME FOR ATTRACTIVE PEOPLE - EPISODE 3","STRONG BAD'S COOL GAME FOR ATTRACTIVE PEOPLE EPISODE 1 HOMESTAR RUINER")
        $gamename = $gamename.Replace("STRONG BAD'S COOL GAME FOR ATTRACTIVE PEOPLE - EPISODE 4","STRONG BAD'S COOL GAME FOR ATTRACTIVE PEOPLE EPISODE 1 HOMESTAR RUINER")
        $gamename = $gamename.Replace("STRONG BAD'S COOL GAME FOR ATTRACTIVE PEOPLE - EPISODE 5","STRONG BAD'S COOL GAME FOR ATTRACTIVE PEOPLE EPISODE 1 HOMESTAR RUINER")
        $gamename = $gamename.Replace("TOKI TORI 2+TOKYO XANADU EX+","TOKYO XANADU EX+")
        $gamename = $gamename.Replace("WARHAMMER: END TIMES - VERMINTIDE","WARHAMMER: END TIMES   VERMINTIDE")
        $gamename = $gamename -replace "^WHITE KNIGHT CHRONICLES$","WHITE KNIGHT CHRONICLES INTERNATIONAL EDITION"
        #REMOVE UNNECESSARY CHARACTERS
        $gamename = $gamename.Replace(" : "," ")
        $gamename = $gamename.Replace(":","")
        $gamename = $gamename.Replace("'","")
        $gamename = $gamename.Replace('`',"")
        $gamename = $gamename.Replace(" - "," ")
        $gamename = $gamename.Replace("-"," ")
        $gamename = $gamename.Replace("_"," ")
        $gamename = $gamename.Replace("/","")
        $gamename = $gamename.Replace("@","")
        $gamename = $gamename.Replace("Ω","OMEGA")
        $gamename = $gamename.Replace(".","")
        $gamename = $gamename.Replace("*","")        
        $gamename = $gamename.Replace(" & "," ")
        $gamename = $gamename.Replace("Û","U")
        $gamename = $gamename.Replace("É","E")
        $gamename = $gamename.Replace("™","")   
        $gamename = $gamename.Replace("®","")
        #FIX TYPOS or UNNECESSARY EDITION VERSIONS
        $gamename = $gamename.Replace(" ULTIMATE EDITION","")
        $gamename = $gamename.Replace(" COMPLETE EDITION","")     
        $gamename = $gamename.Replace(" FULLY GROWN EDITION","")   
        $gamename = $gamename.Replace(" BIG BITE EDITION","")   
        $gamename = $gamename.Replace("OVERKILL EXTENDED CUT","OVERKILL   EXTENDED CUT")   
        $gamename = $gamename.Replace(" WORLD POKER CHAMPIONSHIP","")     
        $gamename = $gamename.Replace("UMBRELLA CHRONICLES","THE UMBRELLA CHRONICLES")   
        $gamename = $gamename.Replace("DISNEY PIXAR ","")        
        $gamename = $gamename.Replace("PANIC!","PANIC")
        $gamename = $gamename.Replace("ZOO!","ZOO")
        $gamename = $gamename.Replace("VOL 1","VOLUME 1")
        $gamename = $gamename.Replace("VOL 2","VOLUME 2")
        $gamename = $gamename.Replace("DISPAIR","DESPAIR")
        $gamename = $gamename.Replace("RIFF ","")    
        $gamename = $gamename.Replace("TREAUSRES","TREASURES")    
        $gamename = $gamename.Replace("NUMENER","NUMENERA")   
        $gamename = $gamename.Replace("UNFINISHED","THE UNFINISHED")        
        $gamename = $gamename.Replace(" BINARY STARS","")       
        $gamename = $gamename.Replace(" FIA WORLD RALLY CHAMPIONSHIP","")        
        $gamename = $gamename.Replace("CHRONOPHANTASMA","CHRONO PHANTASMA")
        $gamename = $gamename.Replace("IN THE MIRACLE WORLD","IN MIRACLE WORLD")
        $gamename = $gamename.Replace(" GAME OF THE YEAR EDITION","")
        $gamename = $gamename.Replace(" GAME OF THE YEAR","")
        $gamename = $gamename.Replace("SLY TRILOGY","SLY COLLECTION")
        $gamename = $gamename -replace "^HOUSE OF THE","THE HOUSE OF THE"
        $gamename = $gamename.Replace("SECOND VELOCITY","SECOND") 
        $gamename = $gamename.Replace("DARKSIDERS 2","DARKSIDERS ii")  
        $gamename = $gamename.Replace("MAFIA 2","MAFIA ii") 
        $gamename = $gamename.Replace("CLASH 2","CLASH ii") 
        $gamename = $gamename.Replace("CHRONICLES 2","CHRONICLES ii") 
        $gamename = $gamename.Replace("WOLFENSTEIN 2","WOLFENSTEIN ii") 
        $gamename = $gamename.Replace("DEAD 3","DEAD III") 
        $gamename = $gamename.Replace("SYBERIA 2","SYBERIA ii") 
        $gamename = $gamename.Replace("KINGDOMS 13","KINGDOMS XIII") 
        $gamename = $gamename.Replace("EPIC MICKEY","EPIC MICKEY 2") 
        $gamename = $gamename.Replace("CARS 2 THE VIDEOGAME", "CARS 2 THE VIDEO GAME")
        $gamename = $gamename.Replace("CRIMES AND PUNISHMENTS", "CRIMES PUNISHMENTS")
        $gamename = $gamename -replace "^STEALTH INC$","STEALTH INC A CLONE IN THE DARK ULTIMATE EDITION"
        $gamename = $gamename.Replace("PILLARS OF ETERNITY","PILLARS OF ETERNITY COMPLETE EDITION")  
        $gamename = $gamename.Replace("PURE HOLDEM","PURE HOLD EM")  
        $gamename = $gamename.Replace("UNDEAD NIGHTMARE","UNDEAD NIGHTMARE PACK") 
        $gamename = $gamename.Replace("INFINITE MINI GOLF","INFINITE MINIGOLF") 
        $gamename = $gamename.Replace("PURE FARMING 18","PURE FARMING 2018") 
        $gamename = $gamename.Replace("WARHAMMER 40000 INQUISITOR MARTYR","WARHAMMER 40000 INQUISITOR - MARTYR")
        #last one, add - instead of spaces
        $gamename = $gamename.Replace(" ","-")
        $gamename = $gamename.ToLower()

        $url = "https://www.metacritic.com/game/"
        if($game.Piattaforma -eq "PS4") {
            $url = "$($url)playstation-4/"
        } elseif ($game.Piattaforma -eq "PS3") {
            $url = "$($url)playstation-3/"
        } elseif ($game.Piattaforma -eq "PS2") {
            $url = "$($url)playstation-2/"
        }
        $url = "$($url)$gamename"
        Write-Host $url
        try
        {
            $response = Invoke-WebRequest -URI $url -ErrorAction Stop
            # This will only execute if the Invoke-WebRequest is successful.
            $StatusCode = $response.StatusCode
        }
        catch
        {
            $StatusCode = $_.Exception.Response.StatusCode.value__
        }
        if($StatusCode -eq "200") {
            Write-Host "SUCCESS"            
            $html = ConvertFrom-Html -Uri $url            
            $metacriticvote = $html.SelectSingleNode("//div[contains(@class, 'details main_details')]//span").InnerText.Trim()
            
            $userscore = $html.SelectSingleNode("//div[contains(@class, 'metascore_w user large game')]")
            if($userscore){
                $userscore = $userscore.InnerText.Trim()                
                $userscore = $userscore.replace(".","")
            } else {
                $userscore = -1
            }
            $genre = $html.SelectSingleNode("//li[contains(@class, 'summary_detail product_genre')]//span[contains(@class, 'data')]").InnerText
            $nplayers = $html.SelectSingleNode("//li[contains(@class, 'summary_detail product_players')]//span[contains(@class, 'data')]").InnerText
        
            if($game.Gioco -eq "DARK MIST") {
                $metacriticvote = -1
                $userscore = -1
                $genre = "Action"
                $nplayers = ""
            }

            if($metacriticvote -eq "No score yet" -or $metacriticvote -eq "tbd"){
                $metacriticvote = -1
            }
            if($userscore -eq "No score yet" -or $userscore -eq "tbd"){
                $userscore = -1
            }
            if(-not $nplayers)
            {
                $nplayers = "N/A"
            }
            if(-not $genre)
            {
                $genre = "N/A"
            }

            Write-Host "$($game.Gioco) - VOTO: $metacriticvote - USERSCORE: $userscore PLAYERS: $nplayers GENERE: $genre"            
            Add-Content -Path $outputFile -Value "$($game.Gioco),$($game.Piattaforma),$($metacriticvote),$userscore,$genre,$nplayers"
        } else {
            Write-HOST "FAILED" -ForegroundColor RED
            Add-Content -Path $outputFile -Value "$($game.Gioco),$($game.Piattaforma),FAILED,,,,"
            Add-Content -Path $errorLogFile -Value $game.Gioco
            Add-Content -Path $errorLogFile -Value $url
        }
    } else {
        Add-Content -Path $outputFile -Value "$($game.Gioco),$($game.Piattaforma),$($game.VotoMetacritic)"
    }    
}