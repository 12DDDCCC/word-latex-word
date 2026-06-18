"""
精简版LaTeX模板提取器 v1.0
只提取生成.tex文件必需的核心参数，不提取编译时模板自动应用的排版细节。
"""
import argparse
from pathlib import Path

from _lite_extractor import TemplateExtractLite
from _skeleton_derive import derive_skeleton_info, derive_metadata_block


__all__ = [
    'TemplateExtractLite',
    'derive_skeleton_info',
    'derive_metadata_block',
    'extract_spec',
]


def extract_spec(cls_path, journal=''):
    """函数级API: 从cls文件提取spec字典

    Args:
        cls_path: .cls文件路径
        journal: 期刊名称（可选，默认从cls文件名推断）

    Returns:
        dict: spec字典（15类核心参数）
    """
    extractor = TemplateExtractLite(cls_path)
    if journal:
        extractor.journal = journal
    return dict(extractor.extract_all())


def main():
    parser = argparse.ArgumentParser(description='精简版LaTeX模板提取器')
    parser.add_argument('--cls-file', required=True, help='模板.cls文件路径')
    parser.add_argument('--journal', default='', help='期刊名称')
    parser.add_argument('--output-dir', default='.', help='输出目录')
    args = parser.parse_args()

    extractor = TemplateExtractLite(args.cls_file)
    if args.journal:
        extractor.journal = args.journal

    out_dir = Path(args.output_dir) / f'{extractor.journal}_lite'
    out_dir.mkdir(parents=True, exist_ok=True)

    spec = extractor.to_json(out_dir / f'{extractor.journal}_template_spec.json')
    extractor.to_guide(out_dir / f'{extractor.journal}_template_guide.md')

    # 统计
    total = len(spec)
    has_data = sum(1 for v in spec.values() if v and v != {} and v != [])
    print(f'提取完成: {has_data}/{total} 类有数据')
    print(f'输出: {out_dir}')


if __name__ == '__main__':
    main()
