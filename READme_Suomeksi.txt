# SAP Salasanan Vaihto Makron Toteutusopas

## Yleiskatsaus
Tässä oppaassa annetaan vaiheittaiset ohjeet SAP:ssä salasanojen vaihtoon automatisoivan VBA (Visual Basic for Applications) makron toteuttamiseen. Makro on suunniteltu toimimaan SAP GUI:n skriptausliittymän kanssa.

## Edellytykset
- Microsoft Excel (mieluiten uusin versio)
- SAP GUI asennettuna tietokoneellesi skriptauksen ollessa sallittu
- Perustason tuntemus Excelistä ja VBA:sta

## Excel-työkirjan Valmistelu
1. **Käytä Makroja Sallivaa Työkirjaa:**
   - Avaa Excel ja luo uusi työkirja.
   - Tallenna työkirja `.xlsm` -päätteellä (Excel Makroja Salliva Työkirja). Tavalliset `.xlsx` työkirjat eivät tue makroja.

2. **Makron Turva-asetukset:**
   - Mene kohtaan `Tiedosto` > `Asetukset` > `Luottamuskeskus` > `Luottamuskeskuksen asetukset`.
   - `Makroasetukset` -osiossa valitse "Poista kaikki makrot käytöstä ilmoituksen kanssa" tai "Salli kaikki makrot". Ensimmäinen vaihtoehto on turvallisempi mutta kysyy aina työkirjaa avattaessa.

3. **Avaa VBA Editori:**
   - Paina `Alt + F11` avataksesi VBA Editorin.

4. **Lisää Uusi Moduuli:**
   - VBA Editorissa, napsauta hiiren kakkospainikkeella `VBAProjekti (SinunTyökirjaNimesi.xlsm)` vasemmassa paneelissa.
   - Valitse `Lisää` > `Moduuli`. Tämä luo uuden moduulin, johon voit liittää makrokoodin.

## Makron Lisääminen
1. **Kopioi Makro Koodi:**
   - Kopioi tarjottu VBA makro koodi.

2. **Liitä Makro Koodi:**
   - VBA Editorissa, liitä kopioitu koodi luomaasi tyhjään moduuliin.

3. **Tallenna Makro:**
   - Paina `Ctrl + S` tallentaaksesi makron työkirjaasi.

## Makron Suorittaminen
1. **Avaa Makroja Salliva Työkirja:**
   - Varmista, että SAP GUI on käynnissä ja kirjautunut sisään.
   - Avaa `.xlsm` työkirja, johon makro on tallennettu.

2. **Suorita Makro:**
   - Voit suorittaa makron monin eri tavoin:
     - Paina `Alt + F8`, valitse makro ja klikkaa "Suorita".
     - Lisää painike Excel-lomakkeellesi, joka laukaisee makron klikattaessa.
     - Kutsu makroa toisesta VBA alirutiinista tai funktiosta.

## Vianetsintä ja Vinkit
- Jos makro ei toimi, varmista että SAP GUI skriptaus on sallittu. Tämä voidaan yleensä asettaa SAP GUI:n asetuksissa "Esteettömyys & Skriptaus" alla.
- Jos saat virheilmoituksia, lue virheviestit huolellisesti. Ne usein antavat vihjeitä siitä, mikä meni pieleen.
- Muista sulkea SAP GUI ja Excel asianmukaisesti makron käytön jälkeen välttääksesi jäljellä olevat prosessit.

## Yhteenveto
Tämä opas tarjoaa perusääriviivat SAP salasanan vaihto makron toteuttamiseen. Muutoksia voi olla tarpeen tehdä riippuen erityisestä SAP-järjestelmästä ja Excel-versiosta. Testaa aina makro ei-tuotantoympäristössä ennen sen käyttöä