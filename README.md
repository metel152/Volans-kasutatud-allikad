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
- smbus2 (MPU6500 / MPU6050 IMU jaoks)
- Adafruit Blinka (board, busio)
- Picamera2
- PySerial
- Matplotlib
- Pillow
- Pandas
- tkintermapview (satelliitkaart)
- PyOpenGL + pyopengltk (3D orientatsioonivaade)
- gspread + google-auth (Google Sheets)
- requests (kaardi tile cache ja HTTP päringud)
- cryptography (PHONE LOC HTTPS sertifikaadi jaoks)
- libcap-dev (Linux süsteemipakett, mitte pip pakett)

---

## Kasutatud andurid

- Raspberry Pi Zero 2W
- BMP280
- IMU 6050 / MPU6500
- RPI Wide NoIR Camera Module 3

---

## Seadistus (Setup)

### 1) Python 3
Raspberry Pi-l peab olema Python 3 installitud (tavaliselt on see juba olemas).

### 2) I2C lubamine (vajalik sensorite jaoks)

Mõlemad sensorid (BMP280 ja MPU6500 / MPU6050) kasutavad I2C protokolli.

- Ava Raspberry Pi terminal  
- Käivita:
  ```bash
  sudo raspi-config
  ```
- Vali:
  ```text
  Interface Options -> I2C -> Enable
  ```
- Taaskäivita süsteem:
  ```bash
  sudo reboot
  ```

### 3) Python teekide installimine

Soovitatav:

```bash
python -m pip install -r requirements.txt
```

Raspberry Pi-s võib vaja minna:

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

   Windowsi maajaama puhul:

   ```powershell
   py maajaam.py
   ```

4. Programmi peatamine:
   ```text
   Ctrl + C
   ```

---

## Maajaama funktsioonid

- Reaalajas telemeetria vastuvõtt LoRa / serial ühenduse kaudu
- Täisekraani dashboard
- Tabidega vaated: satelliitkaart, 3D orientatsioon, graafikud
- 2D attitude indicator
- 3D CanSat orientatsiooni vaade
- Esri satelliitkaardi vaade koos lokaalse tile cache'iga
- Google Maps otsingulingi genereerimine viimasele CanSat asukohale
- Google Sheets live upload
- Eraldi TEST ja REAL andmed
- TEST SIM simulatsioonirežiim
- PC LOC, PHONE LOC, GND=SAT ja AVG GND ground-station asukoha seadmiseks
- HTTPS telefoniasukoha vastuvõtja lokaalse self-signed sertifikaadiga
- Pildi preview
- CSV ja raw log salvestamine

Märkus: Google Earth ja Google Mapsi eraldi integreeritud satelliidivaade on eemaldatud. Rakendus kasutab kaardi jaoks Esri imagery't ning vajadusel loob Google Mapsi otsingulingi viimasele teadaolevale asukohale.

---

## Kaustastruktuur

```text
CanSat_GroundStation_Data/
├─ CSVs/
│  ├─ telemetry.csv
│  ├─ gps.csv
│  ├─ telemetry_TEST.csv
│  └─ gps_TEST.csv
├─ Previews/
├─ Raw_Logs/
├─ Tile_Cache/
└─ Secrets/
   ├─ sheets_config.json
   ├─ sheets_upload_state.json
   ├─ mission_state.json
   ├─ ground_station_location.json
   └─ phone_https/
```

---

## TEST SIM

Rakenduses olev **TEST SIM** nupp simuleerib:

- GPS liikumist
- kõrguse muutust
- temperatuuri ja õhurõhku
- IMU liikumist
- kaardi trajektoori
- graafikute andmeid

TEST andmed:

- salvestatakse eraldi CSV failidesse
- lähevad Google Sheetsis eraldi tabidesse
- märgitakse `source = TEST`
- ei sega päris lennuandmeid

---

## Google Sheets

Google Sheets upload kasutab Service Account JSON faili.

### Seadistus

1. Loo Google Cloudis Service Account
2. Laadi alla `.json` võti
3. Ava JSON fail ja kopeeri `client_email`
4. Jaga Google Sheet sellele emailile Editor õigusega
5. Lisa koodi või kasuta rakenduse salvestatud konfiguratsiooni:

```python
DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/.../edit"
DEFAULT_SERVICE_ACCOUNT_JSON = r"C:\path\to\service-account.json"
```

Sheets tabid:

- `Telemetry`
- `GPS`
- `Telemetry_TEST`
- `GPS_TEST`

Upload kasutab batchimist, et vältida Google Sheets API write quota probleeme.

---

## Asukoha määramine

Ground station asukohta saab määrata mitmel viisil:

- **GND=SAT** – määrab ground stationi viimase CanSat GPS asukoha järgi
- **AVG GND** – võtab mitu GPS näitu ja kasutab keskmist
- **PC LOC** – kasutab Windowsi asukohateenust, kui täpsus on piisav
- **PHONE LOC** – telefon saadab oma GPS asukoha üle hotspot/võrgu
- käsitsi koordinaatide sisestamine

Kui telefonibrauser blokeerib GPS-i, saab kasutada PHONE LOC lehel käsitsi sisestamist.

---

## Google Maps link

Rakendus ei kasuta Google Mapsi tile API-t ega Google Earthi. Selle asemel saab vajutada **GMAPS LINK**, mis:

- kasutab viimast teadaolevat CanSat GPS asukohta
- kui GPS puudub, kasutab ground station asukohta
- avab Google Mapsi otsingulingi
- kopeerib sama lingi clipboardi

Näide lingi formaadist:

```text
https://www.google.com/maps/search/?api=1&query=59.1234567%2C24.1234567
```

---

## Märkused

- Ära jaga Service Account JSON faili avalikult
- TEST SIM ei kirjuta päris CSV failidesse
- Kaart vajab internetti, kuid Esri tile cache vähendab uuesti allalaadimist
- Windows PC LOC võib hotspotiga olla ebatäpne, seega täpsem on GND=SAT või PHONE LOC
- `libcap-dev` ei kuulu requirements.txt faili, sest see installitakse apt kaudu

---

## Meeskond

**Volans**

---

# English Version

## About the Project

We are team Volans, participating in the CanSat Estonia finals.  
This repository contains both the satellite and ground station software used to complete our two missions.

- **Primary mission** – real-time temperature and air pressure measurement with graph visualization
- **Secondary mission** – near-infrared image capture and analysis to distinguish natural and human-made objects

The code is organized into folders based on function.  
Each folder contains a short description.

---

## Built With

- Python 3 (standard library)
- Raspberry Pi
- Adafruit BMP280 library
- smbus2 for the MPU6500 / MPU6050 IMU
- Adafruit Blinka (board, busio)
- Picamera2
- PySerial
- Matplotlib
- Pillow
- Pandas
- tkintermapview for the satellite map
- PyOpenGL + pyopengltk for the 3D orientation view
- gspread + google-auth for Google Sheets
- requests for map tile caching and HTTP requests
- cryptography for the PHONE LOC HTTPS certificate
- libcap-dev as a Linux system package

---

## Sensors and Hardware

- Raspberry Pi Zero 2W
- BMP280
- MPU6050 / MPU6500 IMU
- Raspberry Pi Wide NoIR Camera Module 3

---

## Setup

### 1) Python 3
Python 3 must be installed on the Raspberry Pi. It is usually already installed.

### 2) Enable I2C

Both sensors use the I2C protocol.

```bash
sudo raspi-config
# Interface Options -> I2C -> Enable
sudo reboot
```

### 3) Install Python dependencies

Recommended:

```bash
python -m pip install -r requirements.txt
```

On Raspberry Pi:

```bash
python3 -m pip install -r requirements.txt
```

### 4) Install system packages on Raspberry Pi

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

3. Run the program:

   ```bash
   cd <folder_name>
   python3 <filename>.py
   ```

   On Windows ground station:

   ```powershell
   py maajaam.py
   ```

4. Stop the program:

   ```text
   Ctrl + C
   ```

---

## Ground Station Features

- Real-time telemetry over LoRa / serial
- Fullscreen dashboard
- Tabbed views: satellite map, 3D orientation, charts
- 2D attitude indicator
- 3D CanSat orientation view
- Esri satellite imagery with local tile cache
- Google Maps search-link generation for the last known CanSat location
- Google Sheets live upload
- Separate TEST and REAL data
- TEST SIM simulation mode
- PC LOC, PHONE LOC, GND=SAT and AVG GND for ground-station positioning
- HTTPS phone-location receiver with a local self-signed certificate
- Image preview
- CSV and raw log saving

Note: Google Earth and embedded Google Maps satellite integration have been removed. The app uses Esri imagery for the built-in map and can generate a Google Maps search link for the last known location.

---

## Folder Structure

```text
CanSat_GroundStation_Data/
├─ CSVs/
│  ├─ telemetry.csv
│  ├─ gps.csv
│  ├─ telemetry_TEST.csv
│  └─ gps_TEST.csv
├─ Previews/
├─ Raw_Logs/
├─ Tile_Cache/
└─ Secrets/
   ├─ sheets_config.json
   ├─ sheets_upload_state.json
   ├─ mission_state.json
   ├─ ground_station_location.json
   └─ phone_https/
```

---

## TEST SIM

The **TEST SIM** button simulates:

- GPS movement
- altitude changes
- temperature and pressure
- IMU movement
- map trail
- chart data

TEST data:

- is saved into separate CSV files
- is uploaded to separate Google Sheets tabs
- is marked with `source = TEST`
- does not pollute real flight data

---

## Google Sheets

Google Sheets upload uses a Service Account JSON file.

### Setup

1. Create a Service Account in Google Cloud
2. Download the `.json` key
3. Open the JSON file and copy the `client_email`
4. Share the Google Sheet with that email as Editor
5. Add the config in code or use the saved app configuration:

```python
DEFAULT_GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/.../edit"
DEFAULT_SERVICE_ACCOUNT_JSON = r"C:\path\to\service-account.json"
```

Sheets tabs:

- `Telemetry`
- `GPS`
- `Telemetry_TEST`
- `GPS_TEST`

Uploads are batched to avoid Google Sheets API write-quota errors.

---

## Location Setup

The ground-station position can be set using:

- **GND=SAT** – uses the latest CanSat GPS position
- **AVG GND** – averages multiple GPS fixes
- **PC LOC** – uses Windows Location Service when accuracy is acceptable
- **PHONE LOC** – sends phone GPS position over the hotspot/network
- manual coordinate entry

If the phone browser blocks GPS, use the manual coordinate fields on the PHONE LOC page.

---

## Google Maps Link

The app does not use Google Maps tile API or Google Earth. Instead, **GMAPS LINK**:

- uses the latest known CanSat GPS position
- falls back to the ground-station position if no GPS fix exists
- opens a Google Maps search link
- copies the same link to the clipboard

Example format:

```text
https://www.google.com/maps/search/?api=1&query=59.1234567%2C24.1234567
```

---

## Notes

- Do not share the Service Account JSON file publicly
- TEST SIM does not write into the real CSV files
- The map needs internet, but Esri tile caching reduces repeated downloads
- Windows PC LOC can be inaccurate on mobile hotspots, so GND=SAT or PHONE LOC is usually more accurate
- `libcap-dev` is not included in requirements.txt because it is installed through apt

---

## Team

**Volans**
