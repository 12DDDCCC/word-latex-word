#!/usr/bin/env python3
"""图片顺序匹配工具 - 确认原始高清图片在文档中的插入顺序

核心思路：
1. 从Word文档中提取图片（按文档插入顺序命名fig1, fig2...）
2. 分析提取图片的特征（尺寸、颜色直方图、文件大小）
3. 与原始高清图片库进行匹配
4. 输出映射关系：fig1对应哪张原始图片、fig2对应哪张...

匹配策略（按优先级）：
- 尺寸+文件大小（同一文件）
- 尺寸+颜色直方图相似度
- 尺寸近似+颜色直方图相似度
"""

import os
import re
import sys
from PIL import Image

try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, ValueError):
    pass


def get_image_features(img_path):
    """获取图片特征：尺寸、模式、文件大小、颜色直方图"""
    try:
        img = Image.open(img_path)
        # 缩放到小尺寸计算颜色直方图（用于像素级比较）
        small = img.convert('RGB').resize((64, 64), Image.LANCZOS)
        hist = []
        for i in range(3):
            channel = [p[i] for p in small.getdata()]
            bins = [0] * 16
            for v in channel:
                bins[v // 16] += 1
            hist.extend(bins)
        total = sum(hist)
        if total > 0:
            hist = [h / total for h in hist]
        return {
            'size': img.size,
            'mode': img.mode,
            'file_size': os.path.getsize(img_path),
            'histogram': hist,
        }
    except Exception as e:
        print(f"  警告: 无法读取 {img_path}: {e}")
        return None


def histogram_similarity(hist1, hist2):
    """计算两个直方图的余弦相似度"""
    if hist1 is None or hist2 is None or len(hist1) != len(hist2):
        return 0.0
    import math
    dot = sum(a * b for a, b in zip(hist1, hist2))
    norm1 = math.sqrt(sum(a * a for a in hist1))
    norm2 = math.sqrt(sum(b * b for b in hist2))
    if norm1 == 0 or norm2 == 0:
        return 0.0
    return dot / (norm1 * norm2)


def natural_sort_key(name):
    """Sort file names so fig2 comes before fig10."""
    return [
        int(part) if part.isdigit() else part.lower()
        for part in re.split(r'(\d+)', str(name))
    ]


def match_images(extracted_dir, original_dir):
    """匹配提取的图片和原始高清图片，返回映射关系

    Returns:
        list: [(fig_name, original_name, match_type), ...]
    """
    # 收集提取的图片（按文档插入顺序）
    extracted_files = sorted([
        f for f in os.listdir(extracted_dir)
        if f.lower().startswith('fig') and f.lower().split('.')[-1] in ('jpg', 'jpeg', 'png', 'tiff', 'bmp')
    ], key=natural_sort_key)

    # 收集原始高清图片
    original_files = [
        f for f in os.listdir(original_dir)
        if f.lower().split('.')[-1] in ('jpg', 'jpeg', 'png', 'tiff', 'bmp')
    ]

    print(f"提取的图片: {len(extracted_files)} 张")
    print(f"原始高清图片: {len(original_files)} 张\n")

    # 提取特征
    extracted_features = {}
    for ef in extracted_files:
        epath = os.path.join(extracted_dir, ef)
        features = get_image_features(epath)
        if features:
            extracted_features[ef] = features
            print(f"  {ef}: {features['size']} {features['file_size']//1024}KB")

    original_features = {}
    for of in original_files:
        opath = os.path.join(original_dir, of)
        features = get_image_features(opath)
        if features:
            original_features[of] = features
            print(f"  {of}: {features['size']} {features['file_size']//1024}KB")

    print()

    # 匹配
    matches = []
    used_originals = set()

    for ef, efeat in extracted_features.items():
        candidates = []

        for of, ofeat in original_features.items():
            if of in used_originals:
                continue

            # 策略1：尺寸+文件大小完全相同（同一文件）
            if efeat['size'] == ofeat['size'] and abs(efeat['file_size'] - ofeat['file_size']) < 1000:
                candidates.append((of, 'size+filesize', -2000))
                continue

            # 策略2：尺寸完全匹配 + 直方图相似度
            if efeat['size'] == ofeat['size']:
                sim = histogram_similarity(efeat.get('histogram'), ofeat.get('histogram'))
                candidates.append((of, f'size+hist({sim:.3f})', -1000 + (1 - sim) * 100))
                continue

            # 策略3：尺寸近似匹配（容差50像素）
            w_diff = abs(efeat['size'][0] - ofeat['size'][0])
            h_diff = abs(efeat['size'][1] - ofeat['size'][1])
            if w_diff < 50 and h_diff < 50:
                sim = histogram_similarity(efeat.get('histogram'), ofeat.get('histogram'))
                candidates.append((of, f'approx+hist({sim:.3f})', w_diff + h_diff + (1 - sim) * 100))

        if candidates:
            candidates.sort(key=lambda x: x[2])
            best_match, match_type, _ = candidates[0]
            matches.append((ef, best_match, match_type))
            used_originals.add(best_match)
            print(f"  {ef} -> {best_match} ({match_type})")
        else:
            matches.append((ef, None, 'no_match'))
            print(f"  {ef} -> 未找到匹配")

    return matches


def print_mapping_table(matches, original_dir):
    """打印映射表格"""
    print("\n" + "=" * 80)
    print("图片顺序映射表")
    print("=" * 80)
    print(f"{'文档顺序':<12} {'提取图片':<15} {'原始高清图片':<40} {'匹配方式'}")
    print("-" * 80)

    for i, (fig_name, orig_name, match_type) in enumerate(matches, 1):
        if orig_name:
            # 获取原始图片文件大小
            orig_path = os.path.join(original_dir, orig_name)
            size_kb = os.path.getsize(orig_path) // 1024
            print(f"图{i:<9} {fig_name:<15} {orig_name:<40} {match_type} ({size_kb}KB)")
        else:
            print(f"图{i:<9} {fig_name:<15} {'[未找到匹配]':<40} {match_type}")

    print("-" * 80)
    print(f"总计: {len(matches)} 张图片\n")


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: python image_order_match.py <extracted_dir> <original_dir>")
        print("  extracted_dir: 从Word提取的图片目录（含fig1.jpeg等）")
        print("  original_dir: 原始高清图片目录")
        print()
        print("输出: 文档顺序映射表（fig1对应哪张原始图片）")
        sys.exit(1)

    extracted_dir = sys.argv[1]
    original_dir = sys.argv[2]

    matches = match_images(extracted_dir, original_dir)
    if matches:
        print_mapping_table(matches, original_dir)
