# Volans CanSat Project

## Projektist

Oleme meeskond Volans, mis osaleb CanSat Eesti projekti lõppvoorus.  
Siin on meie satelliidi ja maajaama tarkvara, mida kasutame kahe missiooni edukaks täitmiseks.

- **Esimene missioon** – temperatuuri ja õhurõhu mõõtmine reaalajas ning andmete visualiseerimine graafikutena  
- **Teine missioon** – piltide tegemine lähiinfrapunakaamera abil ning nende analüüs (loodus ja inimtehtud objektide eristamine)

Koodid on sorteeritud kaustadesse vastavalt nende funktsioonile.  
Iga kausta juures on ka lühikirjeldus.

---

## Ehitatud koos

- Python 3 (standard library)
- Raspberry Pi
- Adafruit BMP280 library
- smbus2 (MPU6500 IMU jaoks)
- Adafruit Blinka (board, busio)
- Picamera2
- PySerial
- Matplotlib
- Pillow
- Pandas
- tkintermapview (satelliitkaart)
- PyOpenGL + pyopengltk (3D orientatsioon)
- gspread + google-auth (Google Sheets)
- requests (tile cache ja API päringud)
- libcap-dev (Linux süsteemipakett)

---

## Kasutatud andurid

- Raspberry Pi Zero 2W
- BMP280
- IMU 6050 / MPU6500
- RPI Wide NoIR Camera Module 3

---

## Seadistus (Setup)

### 1) Python 3

Raspberry Pi-l peab olema Python 3 installitud. Tavaliselt on see juba olemas.

### 2) I2C lubamine (vajalik sensorite jaoks)

Mõlemad sensorid (BMP280 ja MPU6500) kasutavad I2C protokolli.

- Ava Raspberry Pi terminal
- Käivita:

```bash
sudo raspi-config
```

- Vali:

```text
Interface Options → I2C → Enable
```

- Taaskäivita süsteem:

```bash
sudo reboot
```

### 3) Python teekide installimine

Soovitatav:

```bash
pip install -r requirements.txt
```

Kui kasutad mitut Pythoni versiooni, siis kasuta parem:

```bash
python -m pip install -r requirements.txt
```

Või Raspberry Pi peal:

```bash
python3 -m pip install -r requirements.txt
```

### 4) Süsteemipaketid (Raspberry Pi jaoks)

```bash
sudo apt install python3-picamera2 python3-smbus i2c-tools libcap-dev
```

---

## Projekti käivitamine (How to Run)

1. Laadi repository GitHubist:

```bash
git clone <repository-link>
cd <projekti-kaust>
```

2. Tee eelnevalt kirjeldatud seadistus (Setup)

3. Käivita programmid terminalis:

```bash
cd <kausta_nimi>
python3 <failinimi>.py
```

Näiteks maajaama käivitamiseks:

```bash
python3 maajaam.py
```

4. Programmi peatamine:

```text
Ctrl + C
```

---

## Maajaama funktsioonid

- Reaalajas telemeetria vastuvõtt LoRa / Serial ühenduse kaudu
- Satelliitkaart Esri imagery ja tile cache toega
- Google Maps Satellite avamine brauseris
- Google Earth live tracking KML failidega
- Google Sheets live upload
- Eraldi TEST ja REAL andmete salvestamine
- Telemeetria graafikud: temperatuur, rõhk ja kõrgus
- 2D attitude indicator
- 3D orientatsiooni vaade
- Pildi eelvaade
- TEST SIM simulatsioonirežiim
- Automaatne kaustade loomine logide, CSV-de, KML-ide ja preview piltide jaoks

---

## Kaustastruktuur

Maajaam loob automaatselt andmete jaoks eraldi peakausta:

```text
CanSat_GroundStation_Data/
├─ CSVs/
│  ├─ telemetry.csv
│  ├─ gps.csv
│  ├─ telemetry_TEST.csv
│  └─ gps_TEST.csv
├─ KMLs/
├─ Previews/
├─ Raw_Logs/
├─ Tile_Cache/
└─ Secrets/
```

Selgitus:

- `CSVs/` – päris ja testandmete CSV failid
- `KMLs/` – Google Earth live tracking failid
- `Previews/` – vastuvõetud pildi eelvaated
- `Raw_Logs/` – toored vastuvõetud paketid
- `Tile_Cache/` – satelliitkaardi cache
- `Secrets/` – lokaalsed seadistusfailid ja olekufailid

---

## TEST SIM

Maajaamas on TEST SIM nupp, millega saab süsteemi testida ilma päris CanSat ühenduseta.

TEST SIM simuleerib:

- GPS liikumist
- kõrguse muutumist
- temperatuuri ja rõhku
- IMU liikumist
- 2D ja 3D orientatsiooni
- kaardi trajektoori
- Google Earth KML positsiooni
- Google Sheets üleslaadimist
- pildi eelvaadet

TEST andmed ei sega päris missiooniandmeid.

Need salvestatakse eraldi failidesse:

```text
telemetry_TEST.csv
gps_TEST.csv
```

Päris andmed salvestatakse failidesse:

```text
telemetry.csv
gps.csv
```

---

## Google Sheets

Google Sheets integratsioon võimaldab telemeetria ja GPS andmeid reaalajas tabelisse laadida.

Sheets failis kasutatakse eraldi töölehti:

- `Telemetry`
- `GPS`
- `Telemetry_TEST`
- `GPS_TEST`

TEST read märgitakse `source = TEST` väärtusega ja neid saab Google Sheetsis eraldi filtreerida.

### Service Account seadistus

1. Loo Google Cloudis Service Account
2. Luba Google Sheets API ja Google Drive API
3. Laadi alla Service Account JSON fail
4. Ava JSON fail ja leia `client_email`
5. Jaga Google Sheet sellele emailile Editor õigusega
6. Lisa koodi Google Sheet link ja JSON faili asukoht:

```python
DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/.../edit"
DEFAULT_SERVICE_ACCOUNT_JSON = r"C:\path\to\service-account.json"
```

Ära jaga Service Account JSON faili avalikult. See sisaldab privaatvõtit.

---

## Google Earth

Maajaam loob Google Earthi jaoks KML failid, mis võimaldavad live trackingut.

Google Earth vaates kuvatakse:

- CanSat hetkepositsioon
- Ground station positsioon
- lennutrajektoor
- kõrguse info

KML failid asuvad kaustas:

```text
CanSat_GroundStation_Data/KMLs/
```

---

## Satelliitkaart

Maajaam kasutab satelliitkaardi jaoks Esri World Imagery kaardikihte.

Kaardil on:

- CanSat marker
- Ground station marker
- lennutrajektoor
- kaugus ja suund ground stationist
- tile cache, mis salvestab kasutatud kaardipildid lokaalselt

---

## Maajaama asukoht

Maajaama asukohta saab määrata mitmel viisil:

- käsitsi koordinaate sisestades
- `GND=SAT` nupuga, kasutades viimast CanSat GPS asukohta
- `AVG GND` nupuga, keskmistades mitu GPS fixi
- `PC LOC` nupuga, kasutades Windows 11 Location Service asukohta

Asukoht salvestatakse lokaalselt, et seda ei peaks iga kord uuesti sisestama.

---

## Märkused

- TEST SIM ei kirjuta päris CSV failidesse
- Service Account JSON faili ei tohi avalikult jagada
- Satelliitkaart vajab internetiühendust, kuid tile cache aitab korduvkasutusel
- Google Sheets API-l on kirjutamispäringute limiidid, seega kasutatakse batch uploadi
- Windowsis ei ole `libcap-dev` vajalik

---

## Meeskond

**Volans**

---

## About the Project

We are team Volans, participating in the CanSat Estonia finals.  
This repository contains both the satellite and ground station software used to complete our two missions.

- **Primary mission** – real-time temperature and air pressure measurement with graph visualization
- **Secondary mission** – near-infrared imaging and analysis to distinguish natural and human-made objects

The code is sorted into folders according to function.  
Each folder includes a short description.

---

## Built With

- Python 3 (standard library)
- Raspberry Pi
- Adafruit BMP280 library
- smbus2 for the MPU6500 IMU
- Adafruit Blinka (board, busio)
- Picamera2
- PySerial
- Matplotlib
- Pillow
- Pandas
- tkintermapview for the satellite map
- PyOpenGL + pyopengltk for the 3D orientation view
- gspread + google-auth for Google Sheets
- requests for tile caching and API requests
- libcap-dev as a Linux system package

---

## Sensors and Hardware

- Raspberry Pi Zero 2W
- BMP280
- IMU 6050 / MPU6500
- RPI Wide NoIR Camera Module 3

---

## Setup

### 1) Python 3

Python 3 must be installed on the Raspberry Pi. It is usually installed by default.

### 2) Enable I2C

Both the BMP280 and MPU6500 use the I2C protocol.

Open the Raspberry Pi terminal and run:

```bash
sudo raspi-config
```

Select:

```text
Interface Options → I2C → Enable
```

Then reboot:

```bash
sudo reboot
```

### 3) Install Python Libraries

Recommended:

```bash
pip install -r requirements.txt
```

If multiple Python versions are installed, use:

```bash
python -m pip install -r requirements.txt
```

On Raspberry Pi, you may prefer:

```bash
python3 -m pip install -r requirements.txt
```

### 4) System Packages for Raspberry Pi

```bash
sudo apt install python3-picamera2 python3-smbus i2c-tools libcap-dev
```

---

## How to Run

1. Clone the repository:

```bash
git clone <repository-link>
cd <project-folder>
```

2. Complete the setup steps above.

3. Run the required program:

```bash
cd <folder_name>
python3 <filename>.py
```

For example, to start the ground station:

```bash
python3 maajaam.py
```

4. Stop the program with:

```text
Ctrl + C
```

---

## Ground Station Features

- Real-time telemetry over LoRa / Serial
- Satellite map using Esri imagery and tile caching
- Google Maps Satellite browser handoff
- Google Earth live tracking using KML
- Google Sheets live upload
- Separate TEST and REAL data logging
- Telemetry charts for temperature, pressure and altitude
- 2D attitude indicator
- 3D orientation view
- Image preview
- TEST SIM simulation mode
- Automatic folder creation for logs, CSV files, KML files and preview images

---

## Folder Structure

The ground station automatically creates a parent data folder:

```text
CanSat_GroundStation_Data/
├─ CSVs/
│  ├─ telemetry.csv
│  ├─ gps.csv
│  ├─ telemetry_TEST.csv
│  └─ gps_TEST.csv
├─ KMLs/
├─ Previews/
├─ Raw_Logs/
├─ Tile_Cache/
└─ Secrets/
```

Description:

- `CSVs/` – real and test CSV files
- `KMLs/` – Google Earth live tracking files
- `Previews/` – received image previews
- `Raw_Logs/` – raw received packets
- `Tile_Cache/` – cached satellite map tiles
- `Secrets/` – local configuration and state files

---

## TEST SIM

The ground station includes a TEST SIM button for testing without a real CanSat connection.

TEST SIM simulates:

- GPS movement
- altitude changes
- temperature and pressure
- IMU movement
- 2D and 3D orientation
- map trajectory
- Google Earth KML position
- Google Sheets upload
- image preview

TEST data does not interfere with real mission data.

It is saved separately:

```text
telemetry_TEST.csv
gps_TEST.csv
```

Real mission data is saved to:

```text
telemetry.csv
gps.csv
```

---

## Google Sheets

The Google Sheets integration uploads telemetry and GPS data in real time.

The spreadsheet uses separate worksheets:

- `Telemetry`
- `GPS`
- `Telemetry_TEST`
- `GPS_TEST`

TEST rows are marked with `source = TEST` and can be filtered separately in Google Sheets.

### Service Account Setup

1. Create a Service Account in Google Cloud
2. Enable the Google Sheets API and Google Drive API
3. Download the Service Account JSON file
4. Open the JSON file and find `client_email`
5. Share the Google Sheet with that email as Editor
6. Add the Google Sheet URL and JSON path in the code:

```python
DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/.../edit"
DEFAULT_SERVICE_ACCOUNT_JSON = r"C:\path\to\service-account.json"
```

Do not share the Service Account JSON file publicly. It contains a private key.

---

## Google Earth

The ground station creates KML files for Google Earth live tracking.

Google Earth displays:

- current CanSat position
- ground station position
- flight trail
- altitude information

KML files are stored in:

```text
CanSat_GroundStation_Data/KMLs/
```

---

## Satellite Map

The ground station uses Esri World Imagery for the satellite map.

The map shows:

- CanSat marker
- ground station marker
- flight trail
- distance and bearing from the ground station
- tile cache for locally storing used map tiles

---

## Ground Station Location

The ground station position can be set in several ways:

- manually by entering coordinates
- with the `GND=SAT` button using the latest CanSat GPS position
- with the `AVG GND` button by averaging multiple GPS fixes
- with the `PC LOC` button using Windows 11 Location Service

The location is saved locally so it does not need to be entered again every time.

---

## Notes

- TEST SIM does not write into the real CSV files
- Do not share the Service Account JSON file
- The satellite map requires internet access, but tile caching helps with repeated use
- Google Sheets API has write request limits, so batch upload is used
- `libcap-dev` is not required on Windows

---

## Team

**Volans**
