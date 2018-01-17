import excel2img
import os

if __name__ == "__main__":
    excelPath = "output_old/excel/"
    pngPath = "output_old/png/"
    for file in os.listdir(excelPath):
        excel2img.export_img(excelPath + file, pngPath + file.replace(".xls", ".png"))

