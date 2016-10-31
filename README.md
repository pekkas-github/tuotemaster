# tuotemaster
<h2>Yleistä</h2>
<p>Tässä repossa on talletettuna tuotemasterin master- ja develop-haarat. Sen tarkoitus on mahdollistaa koodin tuottaminen töissä työkoneella ja kotona kotikoneella. Viimeisin versio välitetään GItHubin kautta.</p>
<h2>Proseduuri</h2>
<h3>Windows + Access -ympäristö</h3>
<p>Koodia kirjoitetaan Accessin IDE-ympäristössä. Lopuksi luokat, modulit ja makrot talletetaan skriptillä gitin workspaceen. Tehty työ vahvistetaan (commit) ja siirretään GitHub:in develop-haaraan (git push origin develop).</p>
<p>Vastaavasti ennen työn aloittamista haetaan viimeisin versio workspaceen (git pull origin develop). Tämän jälkeen VBA-skriptillä siirretään luokat, modulit ja makrot Accessin IDE-ympäristöön. Tiedosto README.mb poistetaan.</p>
<h3>Mac + Smultron</h3>
<p>Ennen työn aloittamista haetaan viimeisin versio workspaceen (git pull origin develop). Muutokset koodataan esim. Smultronilla suoraan workspacessa. Tehty työ vahvistetaan (commit) ja siirretään GitHub:in develop-haaraan (git push origin develop).</p>
<h3>Huomioitavaa</h3>
<p>Lomakkeille kirjoitettua koodia ei pysty tallettamaan tekstitiedostoon ja näin ollen sitä ei voida tuoda GitHubiin</p>
