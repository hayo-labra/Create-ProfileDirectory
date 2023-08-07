<#
.SYNOPSIS
    Generoi profiilikansion, jota voi käyttää harjoitusmateriaalina 
    esimerkiksi PowerShell-aiheisissa tehtävissä. 
.DESCRIPTION
    Skripti generoi kuvitteellisen profiilikansion, joka sisältää 
    Documents, Downloads ja Pictures-kansiot. 
     - Documents-kansio sisältää satunnaisia teksti- ja 
       Word-dokumentteja
     - Downloads-kansio sisältää satunnaisia CSV-tiedostoja ja 
       ladattuja kuvia
     - Pictures-kansio sisältää satunnaisen määrän kuvakansioita.
.EXAMPLE
    PS> .\Create-ProfileDirectory.ps1 -Name kayttaja
.NOTES
    Author: Pekka Tapio Aalto
    Date:   7.8.2023
#>

param (
  # Luotavan profiilikansion nimi
  [string]$Name = "user_" + (Get-Random -Maximum 1000)
)

function Create-Docx {
  <# 
  .SYNOPSIS
      Luo uuden Word-dokumentin ja lisää sisällön rivit
      omina tekstikappaleinaan.    
  .EXAMPLE
      Create-Docx -filename "dokkari.docx" -content "Eka kappale`n`nToka kappale"
  #>
  param (
    # Luotavan Word-dokumentin nimi
    [string]$filename,
    # Dokumentin tekstisisältö 
    [string]$content 
  )

  # Luodaan uusi Word-dokumentti.
  $oWord = New-Object -Com Word.Application
  $oWord.Visible = $true
  $oDoc = $oWord.documents.add()

  # Pilkotaan tekstisisältö rivivaihtojen kohdilta.
  $rows = $content -split "`n`n" 

  # Lisätään tekstirivit omina kappaleinaan.
  foreach ($row in $rows) {
    $oPara = $oDoc.Paragraphs.Add()
    $oPara.Range.Text = $row 
    $oPara.Range.InsertParagraphAfter()
  }
  
  # Tallennetaan ja suljetaan.
  $oDoc.SaveAs($filename)
  $oDoc.Close()
  $oWord.Quit()
}

function Create-Textfile {
  <# 
  .SYNOPSIS
      Luo uuden tekstidokumentin ja lisää sisällön.    
  .EXAMPLE
      Create-Textfile -filename "dokkari.txt" -content "Eka kappale`n`nToka kappale"
  #>  
  param (
    # Luotavan tekstidokumentin nimi
    [string]$filename,
    # Dokumentin tekstisisältö
    [string]$content 
  )
  Set-Content -Path $filename -Value $content
}

# Aloituspäivä 100-199 päivää menneisyydestä, luotavien tiedostojen
# luontiajat alkavat tästä päivästä eteenpäin.
$startdate = (Get-Date).AddDays(-(Get-Random -Minimum 100 -Maximum 200))

# Määritellään luotavien alikansioiden nimet.
Set-Variable dirname_pictures -Option Constant "Pictures"
Set-Variable dirname_documents -Option Constant "Documents"
Set-Variable dirname_downloads -Option Constant "Downloads"

# Selvitetään työkansio.
$workdir = Get-Location

# Määritellään luotavien kansioiden ja tiedostojen lukumäärät.
$num_of_pictures_subfolders = Get-Random -Minimum 10 -Maximum 21
$num_of_pictures_files_min = 5
$num_of_pictures_files_max = 20
$num_of_documents = Get-Random -Minimum 15 -Maximum 30
$num_of_downloaded_stockfiles = Get-Random -Minimum 5 -Maximum 10
$num_of_downloaded_images = Get-Random -Minimum 15 -Maximum 30

# Tulostetaan suoritustilannepalkki.
Write-Progress -Activity "Luodaan kansiorakenne" -PercentComplete 0

# Luodaan profiilikansio.
New-Item -Path $workdir -Name $Name -ItemType "directory"
$path_profile = Join-Path -Path $workdir -ChildPath $Name

# Luodaan Pictures-kuvakansio.
New-Item -Path $path_profile -Name $dirname_pictures -ItemType "directory"
$path_pictures = Join-Path -Path $path_profile -ChildPath $dirname_pictures

# Luodaan Documents-tiedostokansio.
New-Item -Path $path_profile -Name $dirname_documents -ItemType "directory"
$path_documents = Join-Path -Path $path_profile -ChildPath $dirname_documents

# Luodaan Downloads-latauskansio.
New-Item -Path $path_profile -Name $dirname_downloads -ItemType "directory"
$path_downloads = Join-Path -Path $path_profile -ChildPath $dirname_downloads

#----------------------------------------------------
# Pictures-kansio
#----------------------------------------------------

# Tulostetaan suoritustilannepalkki.
Write-Progress -Activity "Luodaan Pictures-kansion sisältö" -PercentComplete 25

# Määritellään tarvittavat muuttujat.
$pic_date = $startdate
$pic_id = Get-Random -Minimum 0 -Maximum 10000

# Luodaan kuvakansioiden ensimmäinen luontipäivä,
# vähintään 3 päivää sitten, maksimissaan 14 päivää sitten.
$dirtime = (Get-Date).AddMinutes((-(Get-Random -Minimum 4320 -Maximum 20160)))

# Luodaan kuvakansiot.
for ($i = 1; $i -le $num_of_pictures_subfolders; $i++) {

  # Tulostetaan suoritustilannepalkki.
  Write-Progress -Activity "Luodaan Pictures-kansion sisältö" -Status "Generoidaan kuvakansiota ($i / $num_of_pictures_subfolders)" -PercentComplete 25

  # Muodostetaan kuvakansion nimi.
  $foldername = $pic_date.ToString("yyyyMMdd")
  $filetime = $pic_date.AddMinutes((Get-Random -Minimum 300 -Maximum 1200))  

  # Luodaan uusi kuvakansio.
  New-Item -Path $path_pictures -Name $foldername -ItemType "directory"

  # Muodostetaan polku kuvakansioon.
  $path_date_folder = Join-Path -Path $path_pictures -ChildPath $foldername
  
  # Arvotaan luotavien kuvien lukumäärä.
  $num_of_pictures = Get-Random -Minimum $num_of_pictures_files_min -Maximum $num_of_pictures_files_max
  
  # Luodaan kuvakansioon kuvat.
  for ($j = 0; $j -le $num_of_pictures; $j++) {

    # Muodostetaan kuvan nimi polun kanssa.
    $filename = Join-Path -Path $path_date_folder -ChildPath ("DSC_" + '{0:d4}' -f $pic_id + ".jpg")

    # Arvotaan kuvan suunta, 66% vaakakuvia, 33% pystykuvia.
    if ((Get-Random -Minimum 1 -Maximum 4) % 3 -eq 0) {
      # modulo 0 = pystykuva
      Invoke-WebRequest https://picsum.photos/1000/1500 -OutFile $filename
    } else {
      # modulo 1 ja 2 = vaakakuva
      Invoke-WebRequest https://picsum.photos/1500/1000 -OutFile $filename
    }

    # Päivitetään kuvan kellonajat.
    $(Get-Item $filename).CreationTime=$($filetime)
    $(Get-Item $filename).LastAccessTime=$($filetime)
    $(Get-Item $filename).LastWriteTime=$($filetime)

    # Kasvatetaan kuvaa eteenpäin ja pyöräytetään numerointi tarvittaessa alusta.
    $pic_id = ($pic_id + 1) % 10000

    # Kasvatetaan kuvan ottohetkeä eteenpäin 2-6 minuuttia.
    $filetime = $filetime.AddMinutes((Get-Random -Minimum 2 -Maximum 7))

  }

  # Päivitetään kuvakansion ajat.
  $(Get-Item $path_date_folder).CreationTime=$($dirtime)
  $(Get-Item $path_date_folder).LastAccessTime=$($dirtime)
  $(Get-Item $path_date_folder).LastWriteTime=$($dirtime)

  # Lisätään päivään 1-6 päivää --> uusi kuvien ottopäivä.
  $pic_date = $pic_date.AddDays((Get-Random -Minimum 1 -Maximum 7))

  # Kasvatetaan kansion aikaa 1-9 minuuttia.
  $dirtime = $dirtime.AddMinutes((Get-Random -Minimum 1 -Maximum 10))
}

#----------------------------------------------------
# Documents-kansio
#----------------------------------------------------

# Tulostetaan suoritustilannepalkki.
Write-Progress -Activity "Luodaan Documents-kansion sisältö" -PercentComplete 50

# Luodaan dokumenttikansion dokumentit. 
For ($i = 1; $i -le $num_of_documents; $i++) {

  # Tulostetaan suoritustilannepalkki.
  Write-Progress -Activity "Luodaan Documents-kansion sisältö" -Status "Generoidaan dokumenttitiedostoa ($i / $num_of_documents)" -PercentComplete 50

  # Dokumentin "luontipäivä" on 1-199 päivää meinneisyydessä + satunnainen määrä sekunteja yhden päivän ajalta.
  $document_date = (Get-Date).AddDays(-(Get-Random -Minimum 1 -Maximum 200)).AddMinutes((Get-Random -Maximum 1440))

  # Arvotaan dokumentissa olevien kappaleiden lukumäärä.
  $paragraphs = Get-Random -Minimum 5 -Maximum 30

  # Generoidaan satunnainen teksti.
  $content = Invoke-WebRequest -URI "https://baconipsum.com/api/?type=meat-and-filler&paras=$paragraphs&format=text"

  # Generoidaan dokumentin satunnaisnimi.
  $filename_start = Invoke-WebRequest -URI "https://random-word-api.vercel.app/api?words=1" | ConvertFrom-Json
  
  # Generoiraan luotavan tiedoston tyyppi.
  #  0 = txt
  #  1 = docx
  $filetype = Get-Random -Maximum 2

  # Luodaan dokumentti.
  Switch ($filetype) {
    "0" {
      $filename = Join-Path -Path $path_documents -ChildPath "$filename_start.txt"
      Create-Textfile -filename $filename -content $content
    }
    "1" {
      $filename = Join-Path -Path $path_documents -ChildPath "$filename_start.docx"
      Create-Docx -filename $filename -content $content
    }
  }

  # Päivitetään tiedoston aikaleimat.
  $(Get-Item $filename).CreationTime=$($document_date)
  $(Get-Item $filename).LastAccessTime=$($document_date)
  $(Get-Item $filename).LastWriteTime=$($document_date)  

}

#----------------------------------------------------
# Downloads-kansio
#----------------------------------------------------

# Tulostetaan suoritustilannepalkki.
Write-Progress -Activity "Luodaan Downloads-kansion sisältö" -PercentComplete 75

# Luodaan taulukko yritysten nimistä ja tunnisteista.
$stock_names = @(
  @("alphabet", "GOOG"),
  @("amazon", "AMZN"),
  @("apple", "AAPL"),
  @("cocacola", "KO"),
  @("costco", "COST"),
  @("crocs", "CROX"),
  @("disney", "DIS"),
  @("dropbox", "DBX"),
  @("gopro", "GPRO"),
  @("microsoft", "MSFT"),
  @("netflix", "NFLX"),
  @("meta", "META"),
  @("nike", "NKE"),
  @("paypal", "PYPL"),
  @("pepsi", "PEP"),
  @("pinterest", "PINS"),
  @("snap", "SNAP"),
  @("sonos", "SONO"),
  @("sony", "SONY"),
  @("spotify", "SPOT"),
  @("starbucks", "SBUX"),
  @("tesla", "TSLA"),
  @("underarmour", "UA"),
  @("walmart", "WMT"),
  @("visa","V")
)

# Sekoitetaan taulukko.
$stock_names = $stock_names | Sort-Object { Get-Random }

# Luodaan ladatut CSV-tiedostot.
For ($i = 1; $i -le $num_of_downloaded_stockfiles; $i++) {

  # Tulostetaan suoritustilannepalkki.
  Write-Progress -Activity "Luodaan Downloads-kansion sisältö" -Status "Generoidaan CSV-tiedostoa ($i / $num_of_downloaded_stockfiles)" -PercentComplete 75

  # Poimitaan listan ensimmäinen alkio.
  $stock_name, $stock_names = $stock_names

  # Osaketietojen "latauspäivä" on 10-199 päivää meinneisyydessä + satunnainen määrä sekunteja yhden päivän ajalta.
  $stock_enddate = (Get-Date).AddDays(-(Get-Random -Minimum 10 -Maximum 200)).AddMinutes((Get-Random -Maximum 1440))

  # Osaketiedot ladataan kuudelta edelliseltä kuukaudelta.
  $stock_starttime = ([DateTimeOffset]$stock_enddate.AddMonths(-6)).ToUnixTimeSeconds()
  $stock_endtime = ([DateTimeOffset]$stock_enddate).ToUnixTimeSeconds()

  # Muodostetaan URI.
  $stockuri = "https://query1.finance.yahoo.com/v7/finance/download/$($stock_name[1])?period1=$stock_starttime&period2=$stock_endtime&interval=1d&events=history&includeAdjustedClose=true"

  # Ladataan tiedot.
  $stockdata = Invoke-WebRequest -URI $stockuri

  # Muodostetaan tiedostonimi ja tallennetaan tiedostoon.
  $filename = Join-Path -Path $path_downloads -ChildPath "$($stock_name[0]).csv"
  Create-Textfile -filename $filename -content $stockdata

  # Päivitetään tiedoston aikaleimat.
  $(Get-Item $filename).CreationTime=$($stock_enddate)
  $(Get-Item $filename).LastAccessTime=$($stock_enddate)
  $(Get-Item $filename).LastWriteTime=$($stock_enddate)

}

# Luodaan ladatut kuvat.
for ($i = 1; $i -le $num_of_downloaded_images; $i++) {

  # Tulostetaan suoritustilannepalkki.
  Write-Progress -Activity "Luodaan Downloads-kansion sisältö" -Status "Ladataan JPG-kuvaa ($i / $num_of_downloaded_images)" -PercentComplete 75

  # Generoidaan kuvan satunnaisnimi.
  $filename_image = Invoke-WebRequest -URI "https://random-word-api.vercel.app/api?words=1" | ConvertFrom-Json
  $filename = Join-Path -Path $path_downloads -ChildPath "$filename_image.jpg"

  # Ladataan kuva.
  Invoke-WebRequest "https://source.unsplash.com/random" -OutFile $filename

  # Kuvan "latauspäivä" on 1-199 päivää meinneisyydessä + satunnainen määrä sekunteja yhden päivän ajalta.
  $imagedate = (Get-Date).AddDays(-(Get-Random -Minimum 1 -Maximum 200)).AddMinutes((Get-Random -Maximum 1440))

  # Päivitetään tiedoston aikaleimat.
  $(Get-Item $filename).CreationTime=$($imagedate)
  $(Get-Item $filename).LastAccessTime=$($imagedate)
  $(Get-Item $filename).LastWriteTime=$($imagedate)

}

# Selvitetään hakemistojen luonti- ja kirjoituspäivät.
$documents_create = (Get-Childitem -Path $path_documents -recurse | Sort-Object -Property LastAccessTime | Select-Object -First 1 LastWriteTime).LastWriteTime
$documents_write = (Get-Childitem -Path $path_documents -recurse | Sort-Object -Property LastAccessTime -Descending | Select-Object -First 1 LastWriteTime).LastWriteTime
$pictures_create = (Get-Childitem -Path $path_pictures -recurse | Sort-Object -Property LastAccessTime | Select-Object -First 1 LastWriteTime).LastWriteTime
$pictures_write = (Get-Childitem -Path $path_pictures -recurse | Sort-Object -Property LastAccessTime -Descending | Select-Object -First 1 LastWriteTime).LastWriteTime
$downloads_create = (Get-Childitem -Path $path_downloads -recurse | Sort-Object -Property LastAccessTime | Select-Object -First 1 LastWriteTime).LastWriteTime
$downloads_write = (Get-Childitem -Path $path_downloads -recurse | Sort-Object -Property LastAccessTime -Descending | Select-Object -First 1 LastWriteTime).LastWriteTime
$profile_create = $startdate
$profile_write = (Get-Childitem -Path $path_profile -recurse | Sort-Object -Property LastAccessTime -Descending | Select-Object -First 1 LastWriteTime).LastWriteTime

# Päivitetään hakemistojen luonti- ja kirjoituspäivät.
$(Get-Item $path_documents).CreationTime=$($documents_create)  
$(Get-Item $path_documents).LastWriteTime=$($documents_write)
$(Get-Item $path_pictures).Creationtime=$($pictures_create)  
$(Get-Item $path_pictures).LastWriteTime=$($pictures_write)
$(Get-Item $path_downloads).Creationtime=$($downloads_create)  
$(Get-Item $path_downloads).LastWriteTime=$($downloads_write)
$(Get-Item $path_profile).CreationTime=$($profile_create)
$(Get-Item $path_profile).LastWriteTime=$($profile_write)

# Tulostetaan lopuksi luotu kansiorakenne.
Clear-Host
Write-Host "----------------------------------------------------------------------------"
Write-Host
Write-Host "Uusi profiilikansio on nyt generoitu!"
Write-Host
tree /f $path_profile
Write-Host
Write-Host "----------------------------------------------------------------------------"