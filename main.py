from PIL import ImageGrab
import win32com.client

FILE_PATH = 'Your path to the file folder'                                  # Example: C:\Users\ExcelSceenshot\data.xlsx
IMAGE_NAME = 'Your image name'                                              # Example: name.png
IMAGE_PATH = 'Your path where you want to save the screenshot'                          
WORKSHEETS_NAME = 'Your sheets name'
SCREEN_AREA = 'The range of cells that you want to save as a screenshot'    # Example: 'A1:D4'

def Screenshot():
    client = win32com.client.Dispatch('Excel.Application')

    wb = client.Workbooks.Open(FILE_PATH)

    wz = wb.Worksheets(WORKSHEETS_NAME)
    wz.Range(SCREEN_AREA).CopyPicture(Format=2)

    save_image_path = f'{IMAGE_PATH}{IMAGE_NAME}'

    img = ImageGrab.grabclipboard()
    img.save(save_image_path)

    client.Quit()


if __name__ == '__main__':
    Screenshot()



