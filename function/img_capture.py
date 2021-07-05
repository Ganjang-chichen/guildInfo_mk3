import pyautogui
import io
import os

key_path = "./key/vision.json"
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = key_path

def img_scanner() :
    chart_location = pyautogui.locateOnScreen('./img/img_scanner/bar.png')
    if chart_location == None :
        return 0
    else :
        chart_size = (chart_location[0], chart_location[1] + 40, chart_location[2], 408)
        chart = pyautogui.screenshot('./img/img_scanner/chart.png', region=chart_size)
        return 1

def get_row(y1, y2) :
    axis_center = (y1 + y2) / 2
    i = 0
    while i < 17 :
        if axis_center >= i * 24 and axis_center < (i + 1) * 24 :
            return i
        i = i + 1

def get_col(x1, x2) :
    col_size = [0, 105, 190, 220, 300, 340, 420, 500]
    axis_center = (x1 + x2) / 2
    i = 0
    while i < 7 :
        if axis_center >= col_size[i] and axis_center < col_size[i + 1] :
            return i
        i = i + 1

def detect_text(path):
    
    list_data = [
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""]
    ]

    """Detects text in the file."""
    from google.cloud import vision
    import io
    client = vision.ImageAnnotatorClient()

    with io.open(path, 'rb') as image_file:
        content = image_file.read()

    image = vision.Image(content=content)

    response = client.text_detection(image=image)
    texts = response.text_annotations

    for text in texts:
        vertices = (['({},{})'.format(vertex.x, vertex.y)
                    for vertex in text.bounding_poly.vertices])

        y1 = int(vertices[0].replace("(", "").replace(")", "").split(",")[1])
        y2 = int(vertices[2].replace("(", "").replace(")", "").split(",")[1])

        if y2 - y1 > 24 :
            continue

        x1 = int(vertices[0].replace("(", "").replace(")", "").split(",")[0])
        x2 = int(vertices[1].replace("(", "").replace(")", "").split(",")[0])

        if x2 - x1 > 105 :
            continue

        row_num = get_row(y1, y2)
        col_num = get_col(x1, x2)
        list_data[row_num][col_num] = list_data[row_num][col_num] + str(text.description).replace(",", "")

    if response.error.message:
        raise Exception(
            '{}\nFor more info on error messages, check: '
            'https://cloud.google.com/apis/design/errors'.format(
                response.error.message))

    for t in list_data :
        if t[4] == '' or t[4] == 'ㅇ' or t[4] == 'o' or t[4] == 'O' :
            t[4] = '0'
        
        if t[5] == '' or t[5] == 'ㅇ' or t[5] == 'o' or t[5] == 'O' :
            t[5] = '0'

        if t[6] == '' or t[6] == 'ㅇ' or t[6] == 'o' or t[6] == 'O' :
            t[6] = '0'

        if t[4] == 'LO' :
            t[4] = 5
    
    return list_data

def init() :
    if os.path.exists(key_path) :
        1
    else :
        return False

    Can_SCAN = img_scanner()
    if Can_SCAN == 0 :
        print("can not scan chart")
        return False
    else :
        list_data = detect_text("./img/img_scanner/chart.png")
    
        return list_data

if __name__ == "__main__" :
    li = init()
    for each in li :
        print(each)

