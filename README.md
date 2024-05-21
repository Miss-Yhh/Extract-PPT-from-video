# EXTRACT PPT

## extract PPT from video : 提取网课视频中的PPT

### 作者：哈尔滨工业大学   张浩  于佳宁  
### 作者邮箱：3142949020@qq.com


> 在桌面标准化环境使用（如windows下），pip大于19.3版本，python大于3.7

### 安装模块

```shell
Pillow，opencv-python，numpy，imagehash,python-pptx
```

## 基础版本
### 从视频中截取视频帧 删除相似度较高的图片 结果保存在图片文件夹frames中
```shell
python main.py 
```

### 转PPT：可人为删除frames文件夹中无关图片 然后转PPT
```shell
python PPT.py 
```

### 参数修改
```shell
video_path：被提取视频的路径
second: 提取视频帧的间隔 可根据视频长度以及演讲者翻页速度酌情修改
threshold：在删除差异帧时，用于判断两张图片在多大程度上可以被认为是相似的
```

## 进阶版本

### 如果您的视频中使用的PPT边界清晰，可以使用下面脚本从视频帧中截取出PPT区域

```shell
python shear.py 
```

这个脚本会遍历文件夹中的所有图片，截取每张图片中可能的PPT部分，然后对比所有的PPT部分，选出最合适的，按照这个比例截取所有的图片。结果存放在Cropped-frames文件夹中，如果此时想要转换成PPT，注意修改脚本中的文件夹路径

```shell
python PPT.py 
```
### 如果已知视频中PPT所在的区域，可以手动输入范围后截取

![alt text](image.png)



