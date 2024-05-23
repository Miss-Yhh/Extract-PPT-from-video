# EXTRACT PPT

## extract PPT from video : 提取网课视频中的PPT

### 作者：HIT-JianingYU   HIT-HaoZhang   
### 作者邮箱：3142949020@qq.com


> 在桌面标准化环境使用（如windows下），pip大于19.3版本，python大于3.7

### 安装模块

```shell
Pillow，opencv-python，numpy，imagehash,python-pptx
```

## 基础版本
### 从视频中截取视频帧 删除相似度较高的图片 结果保存在图片文件夹frames中 然后将留下的图片依次转换为ppt格式
```shell
 python main.py test.mp4 output_presentation.pptx
```
这里的 test.mp4 和 output_presentation.pptx 对应替换为传入视频的路径以及生成PPT的路径


## 进阶版本

### 如果您的视频中使用的PPT边界清晰，可以使用下面脚本从视频帧中截取出PPT区域

```shell
 python shear.py test.mp4 output_presentation.pptx

```

这个脚本会遍历文件夹中的所有图片，截取每张图片中可能的PPT部分，然后对比所有的PPT部分，选出最合适的，按照这个比例截取所有的图片，然后转换为PPT，这里的 test.mp4 和 output_presentation.pptx 对应替换为传入视频的路径以及生成PPT的路径



### 如果截取的视频帧中有一些无关图像 可人为删除frames文件夹中无关图片 然后转PPT
```shell
python ppt.py frames output_presentation.pptx
```
这里的frames为图片文件夹路径，默认为frames，output_presentation.pptx可替换为生成PPT的路径

### 参数（可对应再main函数中修改）
```shell
second: 提取视频帧的间隔 可根据视频长度以及演讲者翻页速度酌情修改
threshold：在删除差异帧时，用于判断两张图片在多大程度上可以被认为是相似的
```

