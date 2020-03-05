### 一、 用FFMPEG 添加一张图片作为微信小视频封面的方法
1、 使用ffmpeg为小视频添加封面，正规的做法是使用命令行，
```
# ffmpeg -i [输入视频] -i [封面图片] -map 0 -map 1 -c copy -c:v:1 png -disposition:v:1 attached_pic [输出视频]，下面是一简单的例子
ffmpeg -i a.mp4 -i fm.png -map 0 -map 1 -c copy -c:v:1 png -disposition:v:1 attached_pic b.mp4
```
这样生成的视频封面在windows的资源管理期里面，缩略图可以看到封面，但是发到微信小程序里面看不到，
而且小程序里面连封面缩略图都不显示，我估计是微信小程序担心人们乱改封面发到微信里面（比如一个广告封面是美女），影响不好。
网上说微信应该是取得视频里面得第一帧作为封面，实验了一下，确实是这样得。

2、 使用ffmpeg为小视频添加微信封面，这个windows都不认，但是微信里面确实是。

- 2.1 把那一张图生成一段非常短得视频，比如0.1秒
```
# ffmpeg -ss 0 -t [时长] -f lavfi -i color=c=0x000000:s=[图片尺寸]:r=[帧率]  -i [输入图片名称] -filter_complex  "[1:v]scale=[输出视频大小][v1];[0:v][v1]overlay=0:0[outv]"  -map [outv] -c:v libx264 [输出名称] -y
ffmpeg -ss 0 -t 0.1 -f lavfi -i color=c=0x000000:s=640x368:r=20  -i fm.png -filter_complex  "[1:v]scale=640:368[v1];[0:v][v1]overlay=0:0[outv]"  -map [outv] -c:v libx264 fm.mp4 -y
```

- 2.2 把这个视频和要添加的视频合起来, 编辑filelist文件，然后执行命令，生成文件d.mp4, 但是没有声音
```
文件 filelist.txt, 内容如下两行
file 'fm.mp4'
file 'a.mp4'

ffmpeg -f concat -i filelist.txt -c copy d.mp4 
```
- 2.3 因为第一个没有配声音，把原来的声音写进去
```
# ffmpeg -i [合并的小视频] -i [原始小视频] -c:v copy -c:a aac -ab [声音的码率] -strict experimental -map 0:v:0 -map 1:a:0 [输出带封面的视频]

ffmpeg -i d.mp4 -i a.mp4 -c:v copy -c:a aac -ab 20480 -strict experimental -map 0:v:0 -map 1:a:0 e.mp4
```

通过以上三步，在微信里面就可以看到这个视频的封面就是我们准备的那张图了，但是在windows的资源管理期里面还是没有设置过来，因为我们本来就没有设置封面，
也不能设置。在微信里面播放的时候，也会闪一下，一般也看不出来