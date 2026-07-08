"""
招投标文件生成系统 —— 主入口
功能：读取招标文件 → 解析要求 → 生成投标文件（DOCX）
"""
import os
import sys
import argparse
from datetime import datetime

from src.bid_document_parser import BiddingDocumentParser
from src.bid_response_generator import BidResponseGenerator
from src.document_output import generate_docx


def main():
    parser = argparse.ArgumentParser(
        description='投标文件自动生成系统 —— 根据招标文件自动生成投标文件 (DOCX)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 基本用法（指定招标文件路径）
  python main.py 招标文件.docx
  
  # 交互式模式（逐步输入信息）
  python main.py --interactive
  
  # 使用LLM生成技术方案
  python main.py 招标文件.docx --api-key YOUR_OPENAI_API_KEY
  
  # 指定输出路径
  python main.py 招标文件.docx --output ./我的投标文件.docx
  
  # 一键生成（从配置文件读取公司信息和报价）
  python main.py 招标文件.docx --config config.json
        """
    )
    
    parser.add_argument('bidding_file', nargs='?', 
                        help='招标文件路径（支持 .pdf, .docx, .txt）')
    parser.add_argument('--interactive', '-i', action='store_true',
                        help='交互式模式，逐步输入投标信息')
    parser.add_argument('--api-key', '-k', type=str, default=None,
                        help='OpenAI API Key，用于AI生成技术方案等内容')
    parser.add_argument('--base-url', type=str, default=None,
                        help='OpenAI API 代理地址（可选）')
    parser.add_argument('--output', '-o', type=str, default=None,
                        help='输出文件路径（默认存放在 output/ 目录）')
    parser.add_argument('--config', '-c', type=str, default=None,
                        help='配置文件路径（JSON格式，含公司信息和报价等）')
    parser.add_argument('--company-profile', type=str, default=None,
                        help='公司信息文件路径（默认使用 data/company_profile.json）')

    args = parser.parse_args()

    print("=" * 60)
    print("  投标文件自动生成系统")
    print("=" * 60)

    # ---- 模式1: 交互式 ----
    if args.interactive:
        run_interactive(args)
        return

    # ---- 模式2: 命令行直接生成 ----
    if not args.bidding_file:
        parser.print_help()
        return

    if not os.path.exists(args.bidding_file):
        print(f"\n错误: 招标文件不存在: {args.bidding_file}")
        sys.exit(1)

    run_auto(args)


def run_auto(args):
    """命令行自动模式"""
    # 1. 解析招标文件
    print(f"\n[1/4] 正在解析招标文件: {args.bidding_file}")
    parser = BiddingDocumentParser()
    bidding_info = parser.parse(args.bidding_file)

    project = bidding_info.get('project_info', {})
    requirements = bidding_info.get('requirements', {})
    
    print(f"  项目名称: {project.get('project_name', '未识别')}")
    print(f"  项目编号: {project.get('project_id', '未识别')}")
    print(f"  招标人:   {project.get('tenderee_name', '未识别')}")
    print(f"  技术要求条目: {len(requirements.get('technical', []))} 条")
    print(f"  要求提交文档: {len(bidding_info.get('required_documents', []))} 项")

    # 2. 加载配置或使用默认
    print(f"\n[2/4] 正在准备投标数据...")
    extra_input = {}
    if args.config:
        import json
        with open(args.config, 'r', encoding='utf-8') as f:
            config = json.load(f)
            extra_input = config.get('bid_info', {})
            print(f"  已从配置文件加载投标准备信息")
    else:
        # 从公司信息中获取默认值
        extra_input['submission_date'] = datetime.now().strftime('%Y年%m月%d日')

    # 3. 生成投标文件内容
    print(f"\n[3/4] 正在生成投标文件内容...")
    company_profile = args.company_profile or None
    generator = BidResponseGenerator(company_profile_path=company_profile)

    # 如果提供了API Key，启用LLM
    if args.api_key:
        generator.enable_llm(args.api_key, args.base_url)
        print(f"  已启用AI智能生成技术方案")

    bid_document = generator.generate(bidding_info, extra_input)
    print(f"  已生成 {len(bid_document)} 个章节")

    # 4. 输出DOCX
    print(f"\n[4/4] 正在生成DOCX文档...")
    # 合并项目信息
    project.update(extra_input)
    
    output_path = generate_docx(bid_document, project, args.output)
    
    print(f"\n{'=' * 60}")
    print(f"  投标文件生成完成！")
    print(f"  输出文件: {output_path}")
    print(f"{'=' * 60}")


def run_interactive(args):
    """交互式模式"""
    print("\n--- 交互式模式 ---\n")
    
    # 输入招标文件路径
    while True:
        bidding_file = input("请输入招标文件路径 (.pdf/.docx/.txt): ").strip()
        if os.path.exists(bidding_file):
            break
        print("文件不存在，请重新输入。")
    
    # 解析
    print("\n正在解析招标文件...")
    parser = BiddingDocumentParser()
    bidding_info = parser.parse(bidding_file)
    
    project = bidding_info.get('project_info', {})
    requirements = bidding_info.get('requirements', {})
    
    print(f"\n已识别项目信息:")
    print(f"  项目名称: {project.get('project_name', '（未识别，将使用占位符）')}")
    print(f"  项目编号: {project.get('project_id', '（未识别）')}")
    print(f"  招标人:   {project.get('tenderee_name', '（未识别）')}")
    print(f"  技术要求: {len(requirements.get('technical', []))} 条")
    print(f"  要求提交: {len(bidding_info.get('required_documents', []))} 项材料")

    # 用户补充信息
    extra_input = {}
    print("\n--- 请补充投标信息（留空使用默认值）---")
    
    if not project.get('project_name'):
        extra_input['project_name'] = input("项目名称: ").strip()
    if not project.get('project_id'):
        extra_input['project_id'] = input("项目编号: ").strip()
    
    bid_amount = input("投标总报价（元）: ").strip()
    if bid_amount:
        extra_input['bid_amount'] = bid_amount
    
    implementation = input("实施周期（如：90天）: ").strip()
    if implementation:
        extra_input['implementation_period'] = implementation
    
    warranty = input("质保期（如：2年）: ").strip()
    if warranty:
        extra_input['warranty_period'] = warranty
    
    extra_input['submission_date'] = datetime.now().strftime('%Y年%m月%d日')

    # LLM选项
    use_llm = input("\n是否使用AI生成技术方案? (y/n, 默认n): ").strip().lower()
    api_key = None
    if use_llm == 'y':
        api_key = input("请输入OpenAI API Key: ").strip()
    
    # 生成
    print("\n正在生成投标文件...")
    generator = BidResponseGenerator()
    if api_key:
        generator.enable_llm(api_key)
    
    bid_document = generator.generate(bidding_info, extra_input)
    
    print("正在生成DOCX文档...")
    output_path = generate_docx(bid_document, {**project, **extra_input}, args.output)
    
    print(f"\n{'=' * 60}")
    print(f"  投标文件生成完成！")
    print(f"  输出文件: {output_path}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
