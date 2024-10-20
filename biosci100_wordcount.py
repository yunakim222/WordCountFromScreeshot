# Credits: all hail the mighty github copilot
from PIL import Image
import io
import win32clipboard
import pytesseract

# Specify the path to the Tesseract executable if it's not in your PATH
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def get_image_from_clipboard():
    win32clipboard.OpenClipboard()
    try:
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
            data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
            bmp_header = (
                b"BM"
                + (len(data) + 14).to_bytes(4, byteorder="little")
                + b"\x00\x00\x00\x00"
                + b"\x36\x00\x00\x00"
                + data
            )
            image = Image.open(io.BytesIO(bmp_header))
            return image
        else:
            raise ValueError("No image data in clipboard")
    finally:
        win32clipboard.CloseClipboard()


def extract_text_from_image(image):
    return pytesseract.image_to_string(image)


def count_words(text):
    words = text.split()
    return len(words)


if __name__ == "__main__":
    try:
        image = get_image_from_clipboard()
        text = extract_text_from_image(image)
        word_count = count_words(text)
        print(f"Word count: {word_count}\n")
        print(f"Extracted text:\n{text}")
    except Exception as e:
        print(f"Error: {e}")
