from pdf2image import convert_from_path
from pdf2image.exceptions import(
    PDFInfoNotInstalledError,
    PDFPageCountError,
    PDFSyntaxError
)
from PIL import Image
import pytesseract


class PDF2Text():
    '''Class to read PDF files using Tesseract OCR and save them as .txt'''
    
    def __init__(self, exe_path):
        '''

        PDF2Text(exe_path)
        
        exe_path (str) : Path to tessaract.exe
        
        See below for more info on Tessaract:
        https://github.com/tesseract-ocr/tesseract
        
        '''
        self.exe_path = exe_path
        self.images = None
        self.texts = None
        
    def open(self, path, prep):
        '''
        Opens and reads a pdf file and stores it as string.

        path (str) : Path to the PDF file for conversion
        prep (bool) : Whether the pages are to be pre-processed
        '''
        #Convert PDF into images, read it using Tesseract OCR
        self.images = convert_from_path(path)
        self.texts = self._ocr(self.images, prep)        
        
    def _ocr(self, images, prep):
        '''
        Reads pages and returns them as a list of strings.

        images (list of PIL Images) : Images of each page to read
        prep (bool) : Whether the image is to be pre-processed
        '''
        #Read each page with pytesseract, return as list
        texts = []
        for image in images:
            pytesseract.pytesseract.tesseract_cmd = self.exe_path
            if prep:
                image = self._pre_process(image)
            text = pytesseract.image_to_string(image)
            texts.append(text)
        return texts

    def _pre_process(self, image):
        '''
        Pre-Process a page for more accurate OCR and returns it.

        image (PIL Image) : Image of a page
        '''
        #Image greyscaling, resizing and thresholding
        return image.convert('L')\
               .resize([3 * i for i in image.size], Image.BICUBIC)\
               .point(lambda val: val > 240 and 255)

    def save(self, path):
        '''
        Saves the OCR results to a file.

        path (str) : Path to the output file.
        '''
        #Join texts and save to file
        output = ''.join(self.texts)
        with open(path, 'w') as outfile:
            outfile.write(output)
