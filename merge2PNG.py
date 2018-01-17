import os
from PIL import Image

UNIT_SIZE = 768  # 高
TARGET_WIDTH = 1184 * 6  # 拼接完后的横向长度

path = "D:/imgs/"
images = []  # 先存储所有的图像的名称
for root, dirs, files in os.walk(path):
    for f in files:
        images.append(f)
for i in range(len(images) / 6):  # 6个图像为一组
    imagefile = []
    j = 0
    for j in range(6):
        imagefile.append(Image.open(path + images[i * 6 + j]))
    target = Image.new('RGB', (TARGET_WIDTH, UNIT_SIZE))
    left = 0
    right = UNIT_SIZE
    for image in imagefile:
        target.paste(image, (left, 0, right, UNIT_SIZE))  # 将image复制到target的指定位置中
        left += UNIT_SIZE  # left是左上角的横坐标，依次递增
        right += UNIT_SIZE  # right是右下的横坐标，依次递增
        quality_value = 100  # quality来指定生成图片的质量，范围是0～100
        target.save(path + '/result/' + os.path.splitext(images[i * 6 + j])[0] + '.png', quality=quality_value)
    imagefile = []
