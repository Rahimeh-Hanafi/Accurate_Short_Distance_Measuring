import serial
import time

def main():
    port = "COM7"   # adjust to the Arduino port
    baudrate = 115200

    arduino = serial.Serial(port, baudrate, timeout=1)
    time.sleep(2)  # wait for Arduino reset

    with open("data.txt", "a") as f:
        count = 0
        while count < 10:
            line = arduino.readline().decode(errors="ignore").strip()
            if not line:
                continue

            # Only keep lines that contain "MeasuredRange" and skip "nan"
            if "MeasuredRange" in line and "nan" not in line:
                if "->" in line:
                    line = line.split("->")[-1].strip()

                print(line)
                f.write(line + "\n")
                f.flush()  # ensure it's saved immediately
                count += 1
        # Write separator after each run
        f.write("---------------------------\n")
        f.flush()

    print("✅ Saved 10 Range readings to data.txt")

if __name__ == "__main__":
    main()
