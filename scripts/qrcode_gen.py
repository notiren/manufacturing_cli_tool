import sys
import subprocess

# Ensure required packages
def ensure_package(pkg, imp=None):
    try:
        __import__(imp or pkg)
    except ImportError:
        print(f"📦 Installing missing package: {pkg}")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", pkg])
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to install {pkg}: {e}")
            sys.exit(1)

ensure_package("qrcode")
ensure_package("Pillow", "PIL")

import qrcode
import os

# Helper for dynamic output folder
def get_output_folder(folder_name="extracted"):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(script_dir).lower() == "scripts":
        base_dir = os.path.dirname(script_dir)
    else:
        base_dir = script_dir
    output_folder = os.path.join(base_dir, folder_name)
    os.makedirs(output_folder, exist_ok=True)
    return output_folder

# Script
def gen_qrcode(data, output_folder):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(data)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")

    base_name = "my_qr_code"
    extension = ".png"
    counter = 1

    filename = f"{base_name}{extension}"
    full_path = os.path.join(output_folder, filename)

    while os.path.exists(full_path):
        filename = f"{base_name}_{counter}{extension}"
        full_path = os.path.join(output_folder, filename)
        counter += 1

    img.save(full_path)
    
    print(f"QR code generated and saved as {full_path}")


def main():
    data = input("Enter the text or URL for your QR code: ")
    output_folder = get_output_folder("extracted")
    gen_qrcode(data, output_folder)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        sys.exit(1)
