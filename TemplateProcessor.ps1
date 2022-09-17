<#
Popis:
Skript vezme csv a sablonu a podle hlavicky v csv nahrazuje jednotlive promenne v dokumentu.

De facto pro kazdy radek v CSV:
    zkopiruj sablonu do noveho souboru ktery se jmenuje Poradi_HodnotaPrvnihoSloupce_vygenerovano.docx
    Pro kazdy sloupec v CSV:
        Nahrad v nove zkopirovanem souboru promennou (nazev sloupce v CSV) za hodnotu z daneho sloupce v CSV
    Uloz soubor

Vstupy:
    $sablona => Sablona dokumentu (docx) ve kterem chceme nahradit zastupne promenne za slova z csv
    $cesta_csv_data => CSV s daty / frazemi, ktere chceme dosadit do dokumentu


#>


$Nastaveni = Get-Content -Path $PSScriptRoot\nastaveni.json | ConvertFrom-Json
$Nastaveni

# Nastaveni funkce pro nahrazeni ve wordu
$MatchCase = $true
$MatchWholeWorld = $true
$MatchWildcards = $false
$MatchSoundsLike = $false
$MatchAllWordForms = $false
$Forward = $false
$Wrap = 1
$Format = $false
$Replace = 2

# promenna s aplikaci word jejiz interni funkce / API pouzivame k nahrazeni v souboru
$Word = New-Object -ComObject Word.Application

# cesta k docx souboru sablony
$Sablona = Get-Item "$($PSScriptRoot)\$($Nastaveni.cesta_sablona)"
$SablonaFiletype = $Sablona.Extension

# CSV s daty zakazniku, -Delimiter => oddelovac
$CsvData = Import-Csv "$($PSScriptRoot)\$($Nastaveni.cesta_csv)" -Delimiter $Nastaveni.oddelovac_csv

# Cas ktery pouzijeme k vytvoreni slozky s vysledky behu tohoto skriptu
$StartCas = Get-Date -Format "yyyy-MM-dd HH-mm-ss"

# Adresar do ktereho zapisujeme vystupy
$VystupAdresar = "$($PSScriptRoot)\$($Nastaveni.vystup_cesta)"

# Podadresar $VystupAdresar - zde vzdy pri kazdem behu vytvorime adresar jehoz nazev je datum a cas behu skriptu
$VysledkyAdresar = $($VystupAdresar + "/" + $StartCas).ToString()


# vytvorime si adresare ktere potrebujeme pro beh, 2>$null znamena ze pripadne errory nepiseme do konzole
mkdir $VystupAdresar 2>$null 

# vytvorime si adresar s datumem / casem behu - nevypisujeme do konzole
# - např. 17.září 2022 13:43:46 bude zapsáno ve formátu 2020-05-17 13-43-46 
mkdir $VysledkyAdresar 1>$null

$PromenneZHlavicky = $CsvData[0].psobject.Properties.name
Write-Output "Promenne k nahrazeni podle CSV"
$PromenneZHlavicky


$CisloRadku = 0

# pro kazdy radek v nasem csv s daty zakazniku
foreach($CsvRadek in $CsvData) {
    $CisloRadku++

    $Prvni = $PromenneZHlavicky[0]
    
    $PrvniCastNazvu = $CsvRadek.$Prvni

    # vygenerujeme nazev souboru, napr. Tomáš Marný_dopis.docx
    $NazevSouboru = "$($CisloRadku)_$($PrvniCastNazvu)_$($Nastaveni.vystup_soubor_suffix)$SablonaFiletype"
    $VyslednySoubor = $($VysledkyAdresar + "\" +  $NazevSouboru)

 
    # prekopirujeme sablonu do noveho souboru do naseho adresare s vysledky ($VysledkyAdreasar)
    Copy-Item -Path $Sablona.FullName.ToString() -Destination $VyslednySoubor

    #Otevreme nasi zkopirovanou sablonu
    $Document = $Word.Documents.Open($VyslednySoubor)
    
    Write-Output ""
    Write-Output "Zpracovavam soubor $VyslednySoubor"
    Write-Output "Nahrazujeme jednotlive promenne"

    foreach($HledanaPromenna in $PromenneZHlavicky) {

        # V dokumentu najdi a nahrad kazdou promennou v sablone za hodnoty z CSV souboru
        # co hledame je prvni parametr, za co nahrazujeme je posledni
        # viz dokumentace => https://learn.microsoft.com/en-us/office/vba/api/word.find.execute

        Write-Output "Nahrazuji promennou $HledanaPromenna za $($CsvRadek.$HledanaPromenna)"
        $Vysledek = $Document.Content.Find.Execute($HledanaPromenna, $MatchCase, $MatchWholeWorld, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $CsvRadek.$HledanaPromenna, $Replace)
        if($Vysledek) {
            Write-Output "Nahrazeno"
        } else {
            Write-Output "Promenna $HledanaPromenna neni v sablone"
        }
    }

    
    $Document.Close(-1) # Zavreme a ulozime, viz https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveoptions
}

$Word.Quit()
$konec = read-host "Zmacknete libovolne tlacitko pro ukonceni..."
