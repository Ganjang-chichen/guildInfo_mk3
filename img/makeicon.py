from PIL import Image

def convert_img_to_ico(filePath, filename) :
    img = Image.open(filePath)
    img.save(filename + ".ico")

convert_img_to_ico("slime-export.png", "slime")
