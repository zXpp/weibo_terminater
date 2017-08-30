# -*- coding: utf-8 -*-
from PIL import Image
import os,glob

def make_thumb(path, thumb_path, size):
	#"""生成缩略
	img = Image.open(path)
	width, height = img.size
	# 裁剪图片成正方形
	if width > height:
		delta = (width - height) / 2
		box = (delta, 0, width - delta, height)
		region = img.crop(box)
	elif height > width:
		delta = (height - width) / 2
		box = (0, delta, width, height - delta)
		region = img.crop(box)
	else:
		region = img

	# 缩放
	thumb = region.resize((size, size), Image.ANTIALIAS)

	base, ext = os.path.splitext(os.path.basename(path))
	filename = os.path.join(thumb_path, '%s_thumb.png' % (base,))
	#print filename
	# 保存
	thumb.save(filename, quality=70)

def png2pngs(filespath,output_file,flag="png"):
	if isinstance(filespath,list):
		files=filespath
	else:
		files = glob.glob(os.path.join(filespath,'*'+u"."+flag))
	"""合并图片"""
	imgs = []
	width, height= 0,0
	# 计算总宽度和长度
	try:
		for file in files:
			img = Image.open(file)
			img = img.convert('RGB') if img.mode != "RGB" else img
			imgs.append(img)
			width = img.size[0] if img.size[0] > width else width
			height += img.size[1]

		# 新建一个白色底的图片
		merge_img = Image.new('RGB', (width, height), 0xffffff)
		cur_height = 0
		for img in imgs:
			# 把图片粘贴上去
			merge_img.paste(img, (0, cur_height))
			cur_height += img.size[1]
		merge_img.save(output_file)#自动识别扩展名
	except Exception,e:
		return e

if __name__ == '__main__':

	#ROOT_PATH = #os.path.abspath(os.path.dirname(__file__))
	IMG_PATH = u"D:\\untar\\Panel\\dist\\haha\\709"#os.path.join(ROOT_PATH, 'img')
	# THUMB_PATH = os.path.join(IMG_PATH, 'thumbs')
	# if not os.path.exists(THUMB_PATH):
	# 	os.makedirs(THUMB_PATH)

	# # 生成缩略图
	# files = glob.glob(os.path.join(IMG_PATH, '*.png'))
	# begin_time = time.clock()
	# for file in files:
	# 	make_thumb(file, THUMB_PATH, 90)
	# end_time = time.clock()
	# print ('make_thumb time:%s' % str(end_time - begin_time))

	# 合并图片
	imgs = glob.glob(os.path.join(IMG_PATH, '*.png'))#abspath:files-list/str
	merge_output = os.path.join(IMG_PATH, 'long.pdf')#absfile
	#begin_time = time.clock()
	png2pngs(imgs, merge_output)
	#end_time = time.clock()
	#print ('merge_thumb time:%s' % str(end_time - begin_time))

