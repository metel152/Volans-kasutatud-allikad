#  See kood on kaamerapiltide logimise süsteem, mis:
#   - juhib Picamera2 kaamerat
#   - teeb kümme pilte sekundis
#   - salvestab pildid määratud kausta
#   - tagab, et pildikaust olemas on
#   - prindib pildifaili asukoha ekraanile
#   - peatub turvaliselt, kui kasutaja vajutab Ctrl+C


import time
from datetime import datetime
import os

#kaust kuhu pildid salvestatakse
kaust = "/home/volans/kaamera-logi"
os.makedirs(kaust, exist_ok=True)

#kaamera nimi
kaamera = Picamera2()

#kindel pildi suurus
config = kaamera.create_still_configuration(
main={"size": (3280, 2464)}
)

#konfigureerib pildi suuruse
kaamera.configure(config)

#loop
try:
    while True:

        kaamera.start()

        nimi = datetime.now().strftime("%Y-%m-%d_%H-%M-%S.jpg")
        asukoht = os.path.join(kaust, nimi)

        kaamera.capture_file(asukoht)

        print("Salvestatud:", asukoht)

        time.sleep(1)

except KeyboardInterrupt:
        print("\nLogimine l6petatud.")

