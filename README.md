# Create-ProfileDirectory.ps1

Skripti generoi profiilikansion, jota voi käyttää harjoitusmateriaalina esimerkiksi PowerShell-aiheisissa tehtävissä. 

## Skriptin asennus

  1. Kopioi tämän hakemiston sisältö haluamaasi sijaintiin, esimerkiksi hakemistoon `C:\profiles`.

  3. Jos et ole vielä antanut PowerShell-skripteille suoritusoikeuksia, niin nyt on siihen hyvä aika. Lisäksi saatat joutua sallimaan ladatun allekirjoittamattoman skriptin suorituksen. Nämä tehdään esimerkiksi seuraavilla komennoilla.

     ```
     PS> Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
     PS> Unblock-File C:\profiles\Create-ProfileDirectory.ps1
     ```

## Skriptin käyttö

  1. Käynnistä skripti komennolla:

     ```
     PS> C:\profiles\Create-ProfileDirectory.ps1 -Name kayttaja
     ```

  2. Odottele, että skripti tekee taikansa ja käytä luotua kansiota miten haluat.

## Käyttöoikeudet

  Tämä skripti hyödyntää seuraavia palveluita:
 
   - [Lorem Picsum](https://picsum.photos/) - kuvakansion kuvat
   - [Bacon Ipsum](https://baconipsum.com/) - dokumenttien tekstit
   - [random-word-api.vercel.app](https://random-word-api.vercel.app/) - dokumenttien ja ladattujen kuvien tiedostonimet
   - [Yahoo Finance](https://finance.yahoo.com/) - ladatut pörssikurssit
   - [Unsplash](https://unsplash.com/) - ladatut kuvat
