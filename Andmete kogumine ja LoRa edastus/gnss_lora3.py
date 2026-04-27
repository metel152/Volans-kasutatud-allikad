from picamera2 import Picamera2
import time
from datetime import datetime
import os
import io
import base64
import board
import busio
import adafruit_bmp280
import smbus2
import serial
import pynmea2
from PIL import Image

# =========================================================
# SEADISTUS
# =========================================================

LORA_PORT = "/dev/ttyUSB0"
LORA_BAUD = 57600

GNSS_PORT = "/dev/ttyS0"
GNSS_BAUD = 115200

ANDURI_KAUST = "/home/volans/andurid-logi"
KAAMERA_KAUST = "/home/volans/kaamera-logi"

BARO_LOGI = os.path.join(ANDURI_KAUST, "baro_logi.csv")
IMU_LOGI = os.path.join(ANDURI_KAUST, "imu_logi.csv")
GNSS_LOGI = os.path.join(ANDURI_KAUST, "gnss_logi.csv")

LOOP_DELAY_S = 1.0
PREVIEW_EVERY_N_IMAGES = 10
PREVIEW_SIZE = (160, 120)
PREVIEW_JPEG_QUALITY = 18
PREVIEW_CHUNK_SIZE = 40

SEA_LEVEL_HPA = 1013.25

MPU_AADRESS = 0x68
MPU_VOOL = 0x6B
KIIR_XOUT_H = 0x3B
GYRO_XOUT_H = 0x43

# =========================================================
# GLOBAALSED SEADMED
# =========================================================

i2c = None
bmp280 = None
bus = None
lora = None
gnss = None
kaamera = None

# =========================================================
# INIT / RECONNECT FUNKTSIOONID
# =========================================================

def init_bmp280():
    try:
        new_i2c = busio.I2C(board.SCL, board.SDA)
        sensor = adafruit_bmp280.Adafruit_BMP280_I2C(new_i2c, address=0x76)
        print("BMP280 ühendatud")
        return new_i2c, sensor
    except Exception as e:
        print("BMP280 init viga:", e)
        return None, None


def init_mpu6500():
    try:
        new_bus = smbus2.SMBus(1)
        new_bus.write_byte_data(MPU_AADRESS, MPU_VOOL, 0)
        print("MPU6500 ühendatud")
        return new_bus
    except Exception as e:
        print("MPU6500 init viga:", e)
        return None


def init_lora():
    try:
        ser = serial.Serial(LORA_PORT, LORA_BAUD, timeout=0.2)
        print(f"LoRa ühendatud pordil {LORA_PORT} @ {LORA_BAUD}")
        return ser
    except Exception as e:
        print("LoRa init viga:", e)
        return None


def init_gnss():
    try:
        ser = serial.Serial(GNSS_PORT, GNSS_BAUD, timeout=0.1)
        print(f"GNSS ühendatud pordil {GNSS_PORT} @ {GNSS_BAUD}")

        configure_galileo_only(ser)
        print("GNSS Galileo-only config uuesti saadetud")

        return ser
    except Exception as e:
        print("GNSS init viga:", e)
        return None


def init_kaamera():
    try:
        cam = Picamera2()
        config = cam.create_still_configuration(main={"size": (3280, 2464)})
        cam.configure(config)
        cam.start()
        time.sleep(2)
        print("Kaamera ühendatud")
        return cam
    except Exception as e:
        print("Kaamera init viga:", e)
        return None


def reconnect_bmp280():
    global i2c, bmp280
    print("Proovin BMP280 uuesti ühendada...")
    i2c, bmp280 = init_bmp280()


def reconnect_mpu6500():
    global bus
    print("Proovin MPU6500 uuesti ühendada...")

    try:
        if bus is not None:
            bus.close()
    except Exception:
        pass

    bus = init_mpu6500()


def reconnect_lora():
    global lora
    print("Proovin LoRa uuesti ühendada...")

    try:
        if lora is not None:
            lora.close()
    except Exception:
        pass

    lora = init_lora()


def reconnect_gnss():
    global gnss
    print("Proovin GNSS uuesti ühendada...")

    try:
        if gnss is not None:
            gnss.close()
    except Exception:
        pass

    gnss = init_gnss()


def reconnect_kaamera():
    global kaamera
    print("Proovin kaamera uuesti käivitada...")

    try:
        if kaamera is not None:
            kaamera.stop()
    except Exception:
        pass

    kaamera = init_kaamera()

# =========================================================
# ABIFUNKTSIOONID
# =========================================================

def arvuta_korgus_m(pressure_hpa, sea_level_hpa=SEA_LEVEL_HPA):
    return 44330.0 * (1.0 - (pressure_hpa / sea_level_hpa) ** 0.1903)


def loe_imu(register):
    if bus is None:
        raise RuntimeError("MPU6500 bus puudub")

    high = bus.read_byte_data(MPU_AADRESS, register)
    low = bus.read_byte_data(MPU_AADRESS, register + 1)
    value = (high << 8) + low

    if value >= 0x8000:
        value -= 65536

    return value


def saada_lora(rida):
    global lora

    if lora is None:
        reconnect_lora()
        return False

    try:
        lora.write((rida + "\n").encode("utf-8"))
        return True
    except Exception as e:
        print("LoRa saatmise viga:", e)
        reconnect_lora()
        return False


def loo_preview_base64(pildi_asukoht):
    with Image.open(pildi_asukoht) as img:
        img = img.convert("L")
        img.thumbnail(PREVIEW_SIZE)
        puhver = io.BytesIO()
        img.save(puhver, format="JPEG", quality=PREVIEW_JPEG_QUALITY)
        return base64.b64encode(puhver.getvalue()).decode("ascii")


def saada_preview_lora(image_seq, timestamp_str, pildi_asukoht):
    try:
        preview_b64 = loo_preview_base64(pildi_asukoht)
        total = (len(preview_b64) + PREVIEW_CHUNK_SIZE - 1) // PREVIEW_CHUNK_SIZE

        saada_lora(f"IMGMETA,{image_seq},{timestamp_str},{os.path.basename(pildi_asukoht)},{total}")

        for n in range(total):
            chunk = preview_b64[n * PREVIEW_CHUNK_SIZE:(n + 1) * PREVIEW_CHUNK_SIZE]
            saada_lora(f"IMG,{image_seq},{n},{total},{chunk}")
            time.sleep(0.03)

        print(f"Preview saadetud: image_seq={image_seq}, chunks={total}")
    except Exception as e:
        print("Preview loomise/saatmise viga:", e)

# =========================================================
# GNSS FUNKTSIOONID
# =========================================================

def ubx_checksum(data):
    ck_a = 0
    ck_b = 0

    for b in data:
        ck_a = (ck_a + b) & 0xFF
        ck_b = (ck_b + ck_a) & 0xFF

    return bytes([ck_a, ck_b])


def send_ubx_valset_bool(ser, settings, save_permanently=False):
    layers = 0x07 if save_permanently else 0x01

    payload = bytearray()
    payload += bytes([0x00, layers, 0x00, 0x00])

    for key_id, value in settings:
        payload += key_id.to_bytes(4, "little")
        payload += bytes([1 if value else 0])

    msg_class = 0x06
    msg_id = 0x8A

    length = len(payload).to_bytes(2, "little")
    body = bytes([msg_class, msg_id]) + length + payload
    packet = b"\xB5\x62" + body + ubx_checksum(body)

    ser.write(packet)
    ser.flush()
    time.sleep(0.2)


def configure_galileo_only(ser):
    CFG_SIGNAL_GPS_ENA  = 0x1031001F
    CFG_SIGNAL_SBAS_ENA = 0x10310020
    CFG_SIGNAL_GAL_ENA  = 0x10310021
    CFG_SIGNAL_BDS_ENA  = 0x10310022
    CFG_SIGNAL_QZSS_ENA = 0x10310024
    CFG_SIGNAL_GLO_ENA  = 0x10310025

    settings = [
        (CFG_SIGNAL_GPS_ENA, 0),
        (CFG_SIGNAL_SBAS_ENA, 0),
        (CFG_SIGNAL_GAL_ENA, 1),
        (CFG_SIGNAL_BDS_ENA, 0),
        (CFG_SIGNAL_QZSS_ENA, 0),
        (CFG_SIGNAL_GLO_ENA, 0),
    ]

    send_ubx_valset_bool(ser, settings, save_permanently=False)
    print("GNSS seadistatud: ainult Galileo RAM-is")


def decimal_to_dms(decimal):
    degrees = int(abs(decimal))
    minutes_full = (abs(decimal) - degrees) * 60
    minutes = int(minutes_full)
    seconds = (minutes_full - minutes) * 60
    return degrees, minutes, seconds


def format_dms_lat(lat):
    suund = "N" if lat >= 0 else "S"
    d, m, s = decimal_to_dms(lat)
    return f"{d}°{m}'{s:.2f}\" {suund}"


def format_dms_lon(lon):
    suund = "E" if lon >= 0 else "W"
    d, m, s = decimal_to_dms(lon)
    return f"{d}°{m}'{s:.2f}\" {suund}"


def loe_gnss():
    global gnss

    if gnss is None:
        reconnect_gnss()
        return None

    viimane_fix = None

    try:
        for _ in range(20):
            raw = gnss.readline()
            if not raw:
                break

            line = raw.decode("ascii", errors="ignore").strip()

            if not line.startswith("$"):
                continue

            try:
                msg = pynmea2.parse(line)
            except pynmea2.ParseError:
                continue

            if msg.sentence_type == "GGA":
                gps_qual = str(getattr(msg, "gps_qual", "0"))

                if gps_qual == "0":
                    continue

                lat = msg.latitude
                lon = msg.longitude
                alt = float(msg.altitude) if msg.altitude not in (None, "") else None
                sats = int(msg.num_sats) if msg.num_sats not in (None, "") else None

                viimane_fix = {
                    "lat_decimal": lat,
                    "lon_decimal": lon,
                    "lat_dms": format_dms_lat(lat),
                    "lon_dms": format_dms_lon(lon),
                    "alt_m": alt,
                    "sats": sats,
                    "gps_qual": gps_qual
                }

    except Exception as e:
        print("GNSS lugemise viga:", e)
        reconnect_gnss()

    return viimane_fix

# =========================================================
# KAUSTAD JA FAILID
# =========================================================

os.makedirs(ANDURI_KAUST, exist_ok=True)
os.makedirs(KAAMERA_KAUST, exist_ok=True)

if not os.path.isfile(BARO_LOGI):
    with open(BARO_LOGI, "w", encoding="utf-8") as f:
        f.write("timestamp,temperature_C,pressure_hPa,altitude_m\n")

if not os.path.isfile(IMU_LOGI):
    with open(IMU_LOGI, "w", encoding="utf-8") as f:
        f.write("timestamp,kiirendus_x,kiirendus_y,kiirendus_z,gyro_x,gyro_y,gyro_z\n")

if not os.path.isfile(GNSS_LOGI):
    with open(GNSS_LOGI, "w", encoding="utf-8") as f:
        f.write("timestamp,latitude_decimal,longitude_decimal,latitude_dms,longitude_dms,altitude_m,satellites,gps_qual\n")

# =========================================================
# SEADMETE KÄIVITAMINE
# =========================================================

i2c, bmp280 = init_bmp280()
bus = init_mpu6500()
lora = init_lora()
gnss = init_gnss()
kaamera = init_kaamera()

print("Login andmeid ja saadan LoRa kaudu. Ctrl+C peatamiseks.")

# =========================================================
# LOOP
# =========================================================

telemetry_seq = 0
image_seq = 0

try:
    while True:
        now_dt = datetime.now()
        now = now_dt.strftime("%Y-%m-%d %H:%M:%S")

        temp = None
        rohk = None
        alt = None

        ax_raw = None
        ay_raw = None
        az_raw = None
        gx_raw = None
        gy_raw = None
        gz_raw = None

        gnss_data = None

        # -------------------------
        # BARO
        # -------------------------
        try:
            if bmp280 is None:
                reconnect_bmp280()

            if bmp280 is not None:
                temp = bmp280.temperature
                rohk = bmp280.pressure
                alt = arvuta_korgus_m(rohk)

                with open(BARO_LOGI, "a", encoding="utf-8") as f:
                    f.write(f"{now},{temp:.2f},{rohk:.2f},{alt:.2f}\n")

                print(f"{now} | BMP280 | {temp:.2f} C | {rohk:.2f} hPa | {alt:.2f} m")
            else:
                print(f"{now} | BMP280 | ühendus puudub")

        except Exception as e:
            print("BMP280 viga:", e)
            reconnect_bmp280()

        # -------------------------
        # IMU
        # -------------------------
        try:
            if bus is None:
                reconnect_mpu6500()

            if bus is not None:
                ax_raw = loe_imu(KIIR_XOUT_H)
                ay_raw = loe_imu(KIIR_XOUT_H + 2)
                az_raw = loe_imu(KIIR_XOUT_H + 4)

                gx_raw = loe_imu(GYRO_XOUT_H)
                gy_raw = loe_imu(GYRO_XOUT_H + 2)
                gz_raw = loe_imu(GYRO_XOUT_H + 4)

                with open(IMU_LOGI, "a", encoding="utf-8") as f:
                    f.write(f"{now},{ax_raw},{ay_raw},{az_raw},{gx_raw},{gy_raw},{gz_raw}\n")

                print(f"{now} | MPU6500 | {ax_raw} | {ay_raw} | {az_raw} | {gx_raw} | {gy_raw} | {gz_raw}")
            else:
                print(f"{now} | MPU6500 | ühendus puudub")

        except Exception as e:
            print("MPU6500 viga:", e)
            reconnect_mpu6500()

        # -------------------------
        # GNSS
        # -------------------------
        try:
            gnss_data = loe_gnss()

            if gnss_data is not None:
                with open(GNSS_LOGI, "a", encoding="utf-8") as f:
                    f.write(
                        f'{now},'
                        f'{gnss_data["lat_decimal"]:.8f},'
                        f'{gnss_data["lon_decimal"]:.8f},'
                        f'"{gnss_data["lat_dms"]}",'
                        f'"{gnss_data["lon_dms"]}",'
                        f'{gnss_data["alt_m"] if gnss_data["alt_m"] is not None else "nan"},'
                        f'{gnss_data["sats"] if gnss_data["sats"] is not None else "nan"},'
                        f'{gnss_data["gps_qual"]}\n'
                    )

                print(
                    f'{now} | GNSS | '
                    f'Lat: {gnss_data["lat_dms"]} ({gnss_data["lat_decimal"]:.8f}) | '
                    f'Lon: {gnss_data["lon_dms"]} ({gnss_data["lon_decimal"]:.8f}) | '
                    f'Alt: {gnss_data["alt_m"]} m | '
                    f'Sats: {gnss_data["sats"]}'
                )
            else:
                print(f"{now} | GNSS | Fix puudub")

        except Exception as e:
            gnss_data = None
            print("GNSS viga:", e)
            reconnect_gnss()

        # -------------------------
        # TELEMETRY LORA
        # -------------------------
        try:
            telemetry_seq += 1

            temp_s = f"{temp:.2f}" if temp is not None else "nan"
            rohk_s = f"{rohk:.2f}" if rohk is not None else "nan"
            alt_s = f"{alt:.2f}" if alt is not None else "nan"

            ax_s = str(ax_raw) if ax_raw is not None else "nan"
            ay_s = str(ay_raw) if ay_raw is not None else "nan"
            az_s = str(az_raw) if az_raw is not None else "nan"
            gx_s = str(gx_raw) if gx_raw is not None else "nan"
            gy_s = str(gy_raw) if gy_raw is not None else "nan"
            gz_s = str(gz_raw) if gz_raw is not None else "nan"

            # Old TEL format stays unchanged for groundstation compatibility.
            telemetry_packet = (
                f"TEL,{telemetry_seq},{now},"
                f"{temp_s},{rohk_s},{alt_s},"
                f"{ax_s},{ay_s},{az_s},{gx_s},{gy_s},{gz_s}"
            )

            saada_lora(telemetry_packet)
            print("LoRa TX:", telemetry_packet)

            # GNSS is sent separately so it does not break TEL parsing.
            if gnss_data is not None:
                gps_alt_s = (
                    f'{gnss_data["alt_m"]:.2f}'
                    if gnss_data["alt_m"] is not None
                    else "nan"
                )

                gps_sats_s = (
                    str(gnss_data["sats"])
                    if gnss_data["sats"] is not None
                    else "nan"
                )

                gps_packet = (
                    f'GPS,{telemetry_seq},{now},'
                    f'{gnss_data["lat_decimal"]:.8f},'
                    f'{gnss_data["lon_decimal"]:.8f},'
                    f'{gps_alt_s},'
                    f'{gps_sats_s},'
                    f'{gnss_data["gps_qual"]}'
                )

                saada_lora(gps_packet)
                print("LoRa TX:", gps_packet)

        except Exception as e:
            print("Telemeetria saatmise viga:", e)
            reconnect_lora()

        # -------------------------
        # KAAMERA
        # -------------------------
        try:
            if kaamera is None:
                reconnect_kaamera()

            if kaamera is not None:
                image_seq += 1
                nimi = now_dt.strftime("%Y-%m-%d_%H-%M-%S.jpg")
                asukoht = os.path.join(KAAMERA_KAUST, nimi)

                kaamera.capture_file(asukoht)
                print("Salvestatud:", asukoht)

                if image_seq % PREVIEW_EVERY_N_IMAGES == 0:
                    saada_preview_lora(image_seq, now_dt.strftime("%Y-%m-%d_%H-%M-%S"), asukoht)
            else:
                print(f"{now} | Kaamera | ühendus puudub")

        except Exception as e:
            print("Kaamera viga:", e)
            reconnect_kaamera()

        time.sleep(LOOP_DELAY_S)

except KeyboardInterrupt:
    print("\nLogimine lõpetatud.")

finally:
    try:
        if kaamera is not None:
            kaamera.stop()
    except Exception:
        pass

    try:
        if lora is not None:
            lora.close()
    except Exception:
        pass

    try:
        if gnss is not None:
            gnss.close()
    except Exception:
        pass

    try:
        if bus is not None:
            bus.close()
    except Exception:
        pass
