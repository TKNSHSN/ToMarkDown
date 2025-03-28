import argparse
import json
import os
import logging
from concurrent import futures
from pathlib import Path
import mammoth
from markdownify import markdownify
import shutil
from urllib.parse import urlparse
import uuid

# 初始化日志配置
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 加载配置文件
try:
    with open('config.json') as f:
        config = json.load(f)
except Exception as e:
    logging.error(f"加载配置文件失败: {str(e)}")
    exit(1)


def process_single_file(file_path):
    try:
        # 准备输出路径
        output_dir = Path(config['output_dir']) / file_path.stem
        output_dir.mkdir(parents=True, exist_ok=True)
        media_dir = output_dir / config['media_subdir']
        media_dir.mkdir(exist_ok=True)

        # 转换文档并提取图片
        with open(file_path, "rb") as f:
            result = mammoth.convert_to_html(f, convert_image=mammoth.images.img_element(lambda img: convert_image(img, media_dir)))
            
        # 保存HTML中间文件
        html_path = output_dir / (file_path.stem + ".html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(result.value)

        # 转换HTML为Markdown
        md_content = markdownify(result.value, heading_style="ATX")
        
        # 保存最终Markdown文件
        md_path = output_dir / (file_path.stem + ".md")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)
            
        logging.info(f"成功转换: {file_path} -> {md_path}")
        
    except Exception as e:
        logging.error(f"处理文件{file_path}失败: {str(e)}")
        raise

# 图片处理回调函数
def convert_image(image, media_dir):
    with image.open() as image_bytes:
        file_ext = image.content_type.split('/')[-1]
        file_name = f"image_{uuid.uuid4().hex[:6]}.{file_ext}"
        save_path = media_dir / file_name
        with open(save_path, "wb") as f:
            shutil.copyfileobj(image_bytes, f)
    return {"src": f"{config['media_subdir']}/{file_name}"}


def main():
    parser = argparse.ArgumentParser(description='Word文档批量转换工具')
    parser.add_argument('input', help='输入文件或目录路径')
    args = parser.parse_args()

    input_path = Path(args.input)
    
    # 验证输出目录
    output_dir = Path(config['output_dir'])
    output_dir.mkdir(parents=True, exist_ok=True)

    # 收集待处理文件
    target_files = []
    if input_path.is_file():
        target_files.append(input_path)
    elif input_path.is_dir():
        for fmt in config['default_formats']:
            target_files.extend(input_path.glob(f'**/*.{fmt}'))
    
    # 并发处理
    with futures.ThreadPoolExecutor(max_workers=config['concurrent_limit']) as executor:
        futures_list = [executor.submit(process_single_file, f) for f in target_files]
        for future in futures.as_completed(futures_list):
            try:
                future.result()
            except Exception as e:
                logging.error(f"处理文件时发生错误: {str(e)}")


if __name__ == '__main__':
    main()