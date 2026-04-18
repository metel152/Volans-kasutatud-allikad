# Volans CanSat Project

## Projektist

Oleme meeskond Volans, mis osaleb CanSat Eesti projekti lõppvoorus.  
Siin on meie satelliidi tarkvara, mida kasutame kahe missiooni edukaks täitmiseks.

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
- libcap-dev

---

## Kasutatud andurid

- Raspberry Pi Zero 2W
- BMP280
- IMU 6050
- RPI Wide NoIR Camera Module 3

---

## Seadistus (Setup)

### 1) Python 3
Raspberry Pi-l peab olema Python 3 installitud (tavaliselt on see juba olemas).

### 2) I2C lubamine (vajalik sensorite jaoks)

Mõlemad sensorid (BMP280 ja MPU6500) kasutavad I2C protokolli.

- Ava Raspberry Pi terminal  
- Käivita:
  sudo raspi-config  
- Vali:
  Interface Options → I2C → Enable  
- Taaskäivita süsteem:
  sudo reboot  

### 3) Python teekide installimine

pip install adafruit-blinka  
pip install pyserial  
pip install matplotlib  
pip install pillow  
pip install pandas  

### 4) Süsteemipaketid (Raspberry Pi jaoks)

sudo apt install python3-picamera2 python3-smbus i2c-tools  

---

## Projekti käivitamine (How to Run)

1. Laadi repository GitHubist:
   git clone <repository-link>  
   cd <projekti-kaust>  

2. Tee eelnevalt kirjeldatud seadistus (Setup)

3. Käivita programmid terminalis:

   cd <kausta_nimi>  
   python3 <failinimi>.py  

4. Programmi peatamine:
   Ctrl + C  

---

## Roadmap

- Automaatsete graafikute genereerimine  
- Piltide analüüsi koodi arendamine  
