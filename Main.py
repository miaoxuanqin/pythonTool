import os

def download_bilibili_video(url):
    command = f'you-get --format=dash-flv-AVC --debug --cookies cookies.txt --output-dir ./videos {url}'
    # command = f'you-get -i --debug --cookies cookies.txt  {url}'
    try:
        os.system(command)
        print(f"视频已下载到 ./videos 文件夹")
    except Exception as e:
        print(f"下载失败: {e}")

video_url = "https://www.bilibili.com/video/BV1ym4y1A7qW"
download_bilibili_video(video_url)